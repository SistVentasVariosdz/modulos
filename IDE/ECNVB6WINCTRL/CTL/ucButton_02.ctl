VERSION 5.00
Begin VB.UserControl ucButton_02 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   Begin VB.PictureBox m_About 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7440
      ScaleHeight     =   2175
      ScaleWidth      =   5655
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6240
      Width           =   5655
   End
End
Attribute VB_Name = "ucButton_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************
'
' Control Name: isButton
'
' Author:       Fred.cpp
'               fred_cpp@gmail.com
'
' Page:         http://mx.geocities.com/fred_cpp/
'
' Current
' Version:      3.7.2
'
' Description:  Multy Style Command Button. Made from scratch,
'               API drawn.
'               I started this project almost 5 years ago, I've
'               learned a lot since then, I think this version
'               Is the first one that can be used In real world,
'               I've keep this control for my own use for a
'               long Time, But I've used It for personal
'               projects, The only one public release was here:
'               http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=51271&lngWId=1
'               I Used It here cuz there was not too much to
'               show then. Afther some months I feel like I
'               could make the final Version, So Here I am.
'               I hope you like It And If you find It useful
'               please vote, leave comments and suggestions,
'               Everything Is wellcome.
'               Best Regards.
'
' Started on:   6-6-2004
'
' 2004-12-6
' Added Soft, Flat, java and Office XP Style
'
' 2004-18-6
' Added Windows XP style
'
' 2004-20-6
' Enhanced Office XP Style
'
' 2004-2-7
' Code Clean Up. Removed Timer, Now Uses the
' Paul Caton's self subclaser
'
' 2004-3-7
' Added Windows Themed Style
'
' 2004-4-7
' Added Galaxy style
'
' 2004-7-7
' Added Enabled/ShowFocus and Custom Colors
'
' 2004-8-7
' Added Plastik Style
'
' 2004-14-7
' Added Keramik Style
'
' 2004-15-7
' Added Multyple styles for Tooltips!
'
' 2004-19-7
' Added Custom WinXP Colors, Disabled State
' for the WinXP style(Normal and custom colors)
'
' 2004-20-7
' Added Mac OSX Style
'
' 2004-24-7
' Keramik Rewritten from scratch
'
' 2004-3-9
' Afther some weeks of inactivity added disabled state for
' plastik and keramik styles
'
' 2004-6-9
' Added a nice About box without a extra form!
' Also some corrections to the MAc OSX Style, But still
' needs more work.
'
' 2004-7-9
' PSC downloads are down. The post was deleted when I tried
' to update. I'm trying to make some improvements to the
' MAc OSX Style, but Nothing too important. just a bugfix
' with the disabled state.
'
' 2004-8-9
' Added font property, also tooltips info is not reseted
' each run. Thanks to the comments of the PSC users I have
' lot of things to improve!
'
' 2004-10-9
' OK, Lot, I mean LOT of suggestions, I would love to be at
' school when I had enought time to code (And I was young
' and was able to sleep one or two hours per day=( )
'
' 2004-12-9
' added font and small offsets for each button style. Also
' Enum names are fixed.
'
' 2004-15-9
' Version 3.0 The last one. added CheckBox Type and value
' property. maybe some Other smal fixes will be made
' soon...maybe.
' Have fun
'
' 2005-27-6
' Version 3.2 The previous was not the last one...
' 3.1 was ActiveX version, this one is usercontrol again.
' A Lot of Bugfixes. Tooltips now are not duplicated (also,
' some memory leaks fixed)
'
' 2005-28-6
' Version 3.3
' Added grayscaled icons for disabled icons
' Also fixed textgrayed when Icon was over the text.
' Thanks to Dennis (dvrdsr)
'
' 2005-1-7
' Version 3.4 Added grayscaled for BMP's when disabled
' (used a Different method). Tooltips bug Fixed forever
' (I really really hope so) Trying to implement Default
' And Cancel properties, but VB not cooperating :)
' some help?
'
' 2005-7-4
' Version: 3.4.1 Fixed Black Icons when using BMP's,
' Removed pMask (Thanks to teee_eeee), Added CustomRegion
' Support by Theme (Thanks to Dana Seaman)
'
' 2005-7-5
' Version 3.4.2 Code Clean Up and Now Working Default and
' Cancel properties, Error Handling in all routines, fixed
' Keyboard events (Done by Aldo Vargas at: 7/4/2005-15:59:11)
'
' 2005-7-5
' Version 3.5.0 Multiline Caption Added. Corrected Java Rects
' Added Custom Offset for text and Icons for each Style
'
' 2005-7-20
' Version 3.5.1 Added RightToLeft property (Still testing)
'
' 2005-7-28
' Version 3.6.0 Fixed CaptionAlign, Added Space Bar Support
' (Thanks to Jiri Novotny)
' Now hover button doesn't bring to front the parent form :)
'
' 2005-7-29
' Version 3.6.1 Code Clean up! RightToLeft fixed?
'
' 2005-9-15
' Version 3.6.2 Fast clicks result in a single click, even
' when the clicks are drawn, no message was fired. fixed.
'
' 2007-12-27
' Version 3.7.0
' Updated Paul Caton Self Subclassed to support Windows 2003 Data Execution Prevention (DEP)
'
' 2008-01-21
' Version 3.7.1
' Empty Caption bug: Fixed
' Trigger the "End" command: Fixed
'
' 2008-10-07
' Version 3.7.2
' Click event fires when button is disabled: Fixed
'
'
Option Explicit
'*************************************************************
'
'   Control Version:
'
Private Const strCurrentVersion = "3.7.2"
'**************************************
'*************************************************************
'
'   Private Constants
'
'**************************************
'Auxiliar Constants
Private Const COLOR_BTNFACE             As Long = 15
Private Const COLOR_BTNSHADOW           As Long = 16
Private Const COLOR_BTNTEXT             As Long = 18
Private Const COLOR_HIGHLIGHT           As Long = 13
Private Const COLOR_WINDOW              As Long = 5
Private Const COLOR_INFOTEXT            As Long = 23
Private Const COLOR_INFOBK              As Long = 24

Private Const BDR_RAISEDOUTER           As Long = &H1
Private Const BDR_SUNKENOUTER           As Long = &H2
Private Const BDR_RAISEDINNER           As Long = &H4
Private Const BDR_SUNKENINNER           As Long = &H8

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT                   As Long = &H1
Private Const BF_TOP                    As Long = &H2
Private Const BF_RIGHT                  As Long = &H4
Private Const BF_BOTTOM                 As Long = &H8
Private Const BF_RECT                   As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

'Windows Messages
Private Const WM_THEMECHANGED           As Long = &H31A
Private Const WM_SYSCOLORCHANGE         As Long = &H15
Private Const WM_USER                   As Long = &H400
Private Const GWL_STYLE                 As Long = -16
Private Const WS_CAPTION                As Long = &HC00000
Private Const WS_THICKFRAME             As Long = &H40000
Private Const WS_MINIMIZEBOX            As Long = &H20000
Private Const SWP_REFRESH               As Long = (&H1 Or &H2 Or &H4 Or &H20)
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_SHOWWINDOW            As Long = &H40
Private Const HWND_TOPMOST              As Long = -&H1
Private Const CW_USEDEFAULT             As Long = &H80000000
'Constants for nPolyFillMode in CreatePolygonRgn y CreatePolyPolygonRgn:
Private Const ALTERNATE                 As Long = 1
''Tooltip Window Constants
Private Const TTS_NOPREFIX              As Long = &H2
'Private Const TTF_TRANSPARENT           As Long = &H100
Private Const TTF_CENTERTIP             As Long = &H2
Private Const TTM_ADDTOOLA              As Long = (WM_USER + 4)
Private Const TTM_DELTOOLA              As Long = (WM_USER + 5)
Private Const TTM_SETTIPBKCOLOR         As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR       As Long = (WM_USER + 20)
Private Const TTM_SETTITLE              As Long = (WM_USER + 32)
Private Const TTM_DELTOOLW              As Long = (WM_USER + 51)
Private Const TTM_ADDTOOLW              As Long = (WM_USER + 50)
Private Const TTM_SETTITLEW             As Long = (WM_USER + 33)
Private Const TTS_BALLOON               As Long = &H40
Private Const TTS_ALWAYSTIP             As Long = &H1
Private Const TTF_SUBCLASS              As Long = &H10
Private Const TOOLTIPS_CLASSA           As String = "tooltips_class32"

'==================================================================================================
'*************************************************************
'
'   Required Type Definitions
'
'*************************************************************
Private Type POINT
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

''Tooltip Window Types
Private Type TOOLINFO
  lSize                           As Long
  lFlags                          As Long
  lHwnd                           As Long
  lId                             As Long
  lpRect                          As RECT
  hInstance                       As Long
  lpStr                           As String
  lParam                          As Long
End Type

Enum isbStyle         'Styles
  [isbNormal] = &H0
  [isbSoft] = &H1
  [isbFlat] = &H2
  [isbJava] = &H3
  [isbOfficeXP] = &H4
  [isbWindowsXP] = &H5
  [isbWindowsTheme] = &H6
  [isbPlastik] = &H7
  [isbGalaxy] = &H8
  [isbKeramik] = &H9
  [isbMacOSX] = &HA
End Enum

Enum isbButtonType
  isbButton = &H0
  isbCheckBox = &H1
End Enum

Enum isbAlign
  [isbCenter] = &H0
  [isbleft] = &H1
  [isbRight] = &H2
  [isbTop] = &H3
  [isbbottom] = &H4
End Enum

Private Enum isState
  statenormal = &H1
  stateHot = &H2
  statePressed = &H3
  statedisabled = &H4
  stateDefaulted = &H5
End Enum

Private Type ICONINFO
  fIcon As Long
  xHotspot As Long
  yHotspot As Long
  hbmMask As Long
  hbmColor As Long
End Type

Private Type BITMAP
  bmType       As Long    'LONG   // Specifies the bitmap type. This member must be zero.
  bmWidth      As Long    'LONG   // Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
  bmHeight     As Long    'LONG   // Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
  bmWidthBytes As Long    'LONG   // Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
  bmPlanes     As Integer 'WORD   // Specifies the count of color planes.
  bmBitsPixel  As Integer 'WORD   // Specifies the number of bits required to indicate the color of a pixel.
  bmBits       As Long    'LPVOID // Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
End Type

'*************************************************************
'
'   Required Enums
'
'*************************************************************

Public Enum ttIconType
  TTNoIcon = 0
  TTIconInfo = 1
  TTIconWarning = 2
  TTIconError = 3
End Enum

Public Enum ttStyleEnum
  TTStandard
  TTBalloon
End Enum

Private Type RGBQUAD
        rgbBlue                     As Byte
        rgbGreen                    As Byte
        rgbRed                      As Byte
        rgbReserved                 As Byte
End Type

Private Type BITMAPINFOHEADER
        biSize                      As Long
        biWidth                     As Long
        biHeight                    As Long
        biPlanes                    As Integer
        biBitCount                  As Integer
        biCompression               As Long
        biSizeImage                 As Long
        biXPelsPerMeter             As Long
        biYPelsPerMeter             As Long
        biClrUsed                   As Long
        biClrImportant              As Long
End Type

Private Type BITMAPINFO
        bmiHeader                   As BITMAPINFOHEADER
        bmiColors                   As RGBQUAD
End Type

Private Enum DrawTextFlags
  DT_TOP = &H0
  DT_LEFT = &H0
  DT_CENTER = &H1
  DT_RIGHT = &H2
  DT_VCENTER = &H4
  DT_BOTTOM = &H8
  DT_WORDBREAK = &H10
  DT_SINGLELINE = &H20
  DT_EXPANDTABS = &H40
  DT_TABSTOP = &H80
  DT_NOCLIP = &H100
  DT_EXTERNALLEADING = &H200
  DT_CALCRECT = &H400
  DT_NOPREFIX = &H800
  DT_INTERNAL = &H1000
  DT_EDITCONTROL = &H2000
  DT_PATH_ELLIPSIS = &H4000
  DT_END_ELLIPSIS = &H8000
  DT_MODIFYSTRING = &H10000
  DT_RTLREADING = &H20000
  DT_WORD_ELLIPSIS = &H40000
  DT_NOFULLWIDTHCHARBREAK = &H80000
  DT_HIDEPREFIX = &H100000
  DT_PREFIXONLY = &H200000
End Enum


Private Type RGBTRIPLE
   rgbBlue  As Byte
   rgbGreen As Byte
   rgbRed   As Byte
End Type

'*************************************************************
'
'   Required API Declarations
'
'*************************************************************
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
                ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, _
                pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemeBackgroundRegion Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, _
                pRect As RECT, pRegion As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
                ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, _
                ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
                ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINT, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, _
                ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, _
                lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, _
                ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long

'*************************************************************
'
'   Private variables
'
'*************************************************************
Private m_bFocused                  As Boolean
Private m_bVisible                   As Boolean
Private m_iState                    As isState
Private m_iStyle                    As isbStyle
Private m_iNonThemeStyle            As isbStyle
Private m_btnRect                   As RECT
Private m_txtRect                   As RECT
Private m_lRegion                   As Long
Private m_sCaption                  As String
Private m_CaptionAlign              As isbAlign
Private m_IconAlign                 As isbAlign
Private m_Icon                      As StdPicture
Private m_Font                      As StdFont
Private m_IconSize                  As Long
Private m_bEnabled                  As Boolean
Private m_bShowFocus                As Boolean
Private m_bUseCustomColors          As Boolean
Private m_lBackColor                As Long
Private m_lHighlightColor           As Long
Private m_lFontColor                As Long
Private m_lFontHighlightColor       As Long
Private m_sToolTipText              As String
Private m_sTooltiptitle             As String
Private m_lToolTipIcon              As ttIconType
Private m_lToolTipType              As ttStyleEnum
Private m_lttBackColor              As Long
Private m_lttForeColor              As Long
Private m_lttCentered               As Boolean
Private m_lttHwnd                   As Long
Private m_ButtonType                As isbButtonType
Private m_Value                     As Boolean
Private m_MaskColor                 As Long
Private m_UseMaskColor              As Boolean
Private m_bRoundedBordersByTheme    As Boolean
Private m_bRTLText                   As Long
Private lPrevStyle                  As Long
Private iStyleIconOffset            As Long

'Auxiliar Variables
Dim lwFontAlign                     As Long
Dim lPrevButton                     As Long
Dim ttip                            As TOOLINFO

'*************************************************************
'
'   SUB CLASSES
'
'*************************************************************

Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2                                     'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of the User-defined callback parameter data index

Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Const WM_MOUSEMOVE    As Long = &H200
Private Const WM_MOUSELEAVE   As Long = &H2A3
Private Const WM_MOVING       As Long = &H216
Private Const WM_SIZING       As Long = &H214
Private Const WM_EXITSIZEMOVE As Long = &H232

Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                      As Long
  dwFlags                     As TRACKMOUSEEVENT_FLAGS
  hwndTrack                   As Long
  dwHoverTime                 As Long
End Type

Private bTrack                As Boolean
Private bTrackUser32          As Boolean
Private bInCtrl               As Boolean
Private bMoving               As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'*************************************************************
'
'   Public Events
'
'*************************************************************
Public Event Click()
Public Event MouseEnter()
Public Event MouseLeave()

' Paul Caton Self Subclassed template
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'======================================================================================================

'-SelfSub code------------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = False) As Boolean 'Subclass the specified window handle
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If

  nMyID = GetCurrentProcessId                                               'Get this process's ID
  GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
  If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
  If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
    
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
  End If
  
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

  If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection
    On Error GoTo 0
  
    If bIdeSafety Then                                                      'If the user wants IDE protection
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    End If
    
    z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
    z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
    z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
    z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
    sc_Subclass = True                                                      'Indicate success
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  
  Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim I As Long

  If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started
    With z_Funk
      For I = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        z_ScMem = .Item(I)                                                  'Get the thunk address
        If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next I                                                                'Next member of the collection
    End With
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
      zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages
      zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages
    End If
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
      zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
    End If
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
      zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
      zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
    End If
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
      zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
    End If
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
  End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim I      As Long                                                        'Loop index

  nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                    'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count
    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If

    For I = 1 To nCount                                                     'Loop through the table entries
      If zData(I) = 0 Then                                                  'If the element is free...
        zData(I) = uMsg                                                     'Use this element
        GoTo Bail                                                           'Bail
      ElseIf zData(I) = uMsg Then                                           'If the message is already in the table...
        GoTo Bail                                                           'Bail
      End If
    Next I                                                                  'Next message table entry

    nCount = I                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim I      As Long                                                        'Loop index

  nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For I = 1 To nCount                                                     'Loop through the table entries
      If zData(I) = uMsg Then                                               'If the message is found...
        zData(I) = 0                                                        'Null the msg value -- also frees the element for re-use
        GoTo Bail                                                           'Bail
      End If
    Next I                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  'MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address
    zMap_hWnd = z_ScMem
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, I, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, I, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, I, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  I = I + 4                                                                 'Bump to the next entry
  j = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While I < j
    RtlMoveMemory VarPtr(nAddr), I, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    I = I + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    FreeLibrary hMod
  End If
End Function

'*************************************************************
'
'added by teee_eeee: unneded pMask Picture Box
'
'*************************************************************

Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False, Optional ByVal XPBlend As Boolean = False)

   If DstW = 0 Or DstH = 0 Then Exit Sub

   Dim B        As Long
   Dim H        As Long
   Dim F        As Long
   Dim I        As Long
   Dim newW     As Long
   Dim TmpDC    As Long
   Dim TmpBmp   As Long
   Dim TmpObj   As Long
   Dim Sr2DC    As Long
   Dim Sr2Bmp   As Long
   Dim Sr2Obj   As Long
   Dim Data1()  As RGBTRIPLE
   Dim Data2()  As RGBTRIPLE
   Dim Info     As BITMAPINFO
   Dim BrushRGB As RGBTRIPLE
   Dim gCol     As Long

   Dim SrcDC    As Long
   Dim tObj     As Long
   Dim ttt      As Long

   SrcDC = CreateCompatibleDC(hdc)

   If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
   If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)

   If SrcPic.Type = 1 Then
      tObj = SelectObject(SrcDC, SrcPic)
   Else
      Dim hBrush As Long
      tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
      hBrush = CreateSolidBrush(MaskColor)
      DrawIconEx SrcDC, 0, 0, SrcPic.handle, 0, 0, 0, hBrush, &H1 Or &H2
      DeleteObject hBrush
   End If

   TmpDC = CreateCompatibleDC(SrcDC)
   Sr2DC = CreateCompatibleDC(SrcDC)
   TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   TmpObj = SelectObject(TmpDC, TmpBmp)
   Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
   
   ReDim Data1(DstW * DstH * 3 - 1)
   ReDim Data2(UBound(Data1))
   
   With Info.bmiHeader
      .biSize = Len(Info.bmiHeader)
      .biWidth = DstW
      .biHeight = DstH
      .biPlanes = 1
      .biBitCount = 24
   End With

   BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
   BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
   GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
   GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0

   If BrushColor > 0 Then
      BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
      BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
      BrushRGB.rgbRed = BrushColor And &HFF
   End If

   If Not m_UseMaskColor Then TransColor = -1

   newW = DstW - 1

   For H = 0 To DstH - 1
      F = H * DstW
      For B = 0 To newW
         I = F + B
         If GetNearestColor(hdc, CLng(Data2(I).rgbRed) + 256& * Data2(I).rgbGreen + 65536 * Data2(I).rgbBlue) <> TransColor Then
            With Data1(I)
               If BrushColor > -1 Then
                  If MonoMask Then
                     If (CLng(Data2(I).rgbRed) + Data2(I).rgbGreen + Data2(I).rgbBlue) <= 384 Then Data1(I) = BrushRGB
                  Else
                     Data1(I) = BrushRGB
                  End If
               Else
                  If isGreyscale Then
                     gCol = CLng(Data2(I).rgbRed * 0.3) + Data2(I).rgbGreen * 0.59 + Data2(I).rgbBlue * 0.11
                     .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                  Else
                     If XPBlend Then
                        .rgbRed = (CLng(.rgbRed) + Data2(I).rgbRed * 2) \ 3
                        .rgbGreen = (CLng(.rgbGreen) + Data2(I).rgbGreen * 2) \ 3
                        .rgbBlue = (CLng(.rgbBlue) + Data2(I).rgbBlue * 2) \ 3
                     Else
                        Data1(I) = Data2(I)
                     End If
                  End If
               End If
            End With
            
         End If
      
      Next B
   
   Next H

   SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0

   Erase Data1, Data2
   DeleteObject SelectObject(TmpDC, TmpObj)
   DeleteObject SelectObject(Sr2DC, Sr2Obj)
   If SrcPic.Type = 3 Then DeleteObject SelectObject(SrcDC, tObj)
   DeleteDC TmpDC: DeleteDC Sr2DC
   DeleteObject tObj: DeleteDC SrcDC

End Sub

'*************************************************************
'
'added by Dennis (dvrdsr) function excerpted vlad memdc class'
'
'*************************************************************
Public Function PaintIconGrayscale(ByVal Dest_hDC As Long, _
                                   ByVal hIcon As Long, _
                                   Optional ByVal Dest_X As Long, _
                                   Optional ByVal Dest_Y As Long, _
                                   Optional ByVal Dest_Height As Long, _
                                   Optional ByVal Dest_Width As Long) As Boolean
  On Error GoTo PaintIconGrayscale_Error

  Dim hBMP_Mask  As Long
  Dim hBMP_Image As Long
  Dim hBMP_Prev  As Long
  Dim hIcon_Temp As Long
  Dim hDC_Temp   As Long

  ' Make sure parameters passed are valid
  If Dest_hDC = 0 Or hIcon = 0 Then Exit Function

  ' Extract the bitmaps from the icon
  If pvGetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False Then Exit Function
  ' Create a memory DC to work with
  hDC_Temp = CreateCompatibleDC(0)

  If hDC_Temp = 0 Then GoTo CleanUp

  ' Make the image bitmap gradient
  If pvRenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0) = False Then GoTo CleanUp
  ' Extract the gradient bitmap out of the DC
  SelectObject hDC_Temp, hBMP_Prev
  ' Take the newly gradient bitmap and make a gradient icon from it
  hIcon_Temp = pvCreateIconFromBMP(hBMP_Mask, hBMP_Image)

  If hIcon_Temp = 0 Then GoTo CleanUp

  ' Draw the newly created gradient icon onto the specified DC
  If DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, &H3) <> 0 Then
    PaintIconGrayscale = True
  End If

CleanUp:
  DestroyIcon hIcon_Temp: hIcon_Temp = 0
  DeleteDC hDC_Temp: hDC_Temp = 0
  DeleteObject hBMP_Mask: hBMP_Mask = 0
  DeleteObject hBMP_Image: hBMP_Image = 0
  Exit Function

PaintIconGrayscale_Error:
End Function

Private Function pvGetIconBitmaps(ByVal hIcon As Long, _
                                  ByRef Return_hBmpMask As Long, _
                                  ByRef Return_hBmpImage As Long) As Boolean
  On Error GoTo pvGetIconBitmaps_Error

  Dim TempICONINFO As ICONINFO

  If GetIconInfo(hIcon, TempICONINFO) = 0 Then Exit Function
  Return_hBmpMask = TempICONINFO.hbmMask
  Return_hBmpImage = TempICONINFO.hbmColor
  pvGetIconBitmaps = True
  Exit Function

pvGetIconBitmaps_Error:
End Function

Private Function pvRenderBitmapGrayscale(ByVal Dest_hDC As Long, _
                                         ByVal hBitmap As Long, _
                                         Optional ByVal Dest_X As Long, _
                                         Optional ByVal Dest_Y As Long, _
                                         Optional ByVal Srce_X As Long, _
                                         Optional ByVal Srce_Y As Long) As Boolean
  On Error GoTo pvRenderBitmapGrayscale_Error

  Dim TempBITMAP  As BITMAP
  Dim hScreen     As Long
  Dim hDC_Temp    As Long
  Dim hBMP_Prev   As Long
  Dim MyCounterX  As Long
  Dim MyCounterY  As Long
  Dim NewColor    As Long
  Dim hNewPicture As Long
  Dim DeletePic   As Boolean

  ' Make sure parameters passed are valid
  If Dest_hDC = 0 Or hBitmap = 0 Then Exit Function
  ' Get the handle to the screen DC
  hScreen = GetDC(0)

  If hScreen = 0 Then Exit Function
  ' Create a memory DC to work with the picture
  hDC_Temp = CreateCompatibleDC(hScreen)

  If hDC_Temp = 0 Then GoTo CleanUp
  ' If the user specifies NOT to alter the original, then make a copy of it to use
  DeletePic = False
  hNewPicture = hBitmap
  ' Select the bitmap into the DC
  hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)

  ' Get the height / width of the bitmap in pixels
  If GetObjectAPI(hNewPicture, Len(TempBITMAP), TempBITMAP) = 0 Then GoTo CleanUp
  If TempBITMAP.bmHeight <= 0 Or TempBITMAP.bmWidth <= 0 Then GoTo CleanUp

  ' Loop through each pixel and conver it to it's grayscale equivelant
  For MyCounterX = 0 To TempBITMAP.bmWidth - 1
    For MyCounterY = 0 To TempBITMAP.bmHeight - 1
      NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)

      If NewColor <> -1 Then

        Select Case NewColor
            ' If the color is already a grey shade, no need to convert it
          Case vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
            NewColor = NewColor
          Case Else
            NewColor = 0.33 * (NewColor Mod 256) + 0.59 * ((NewColor \ 256) Mod 256) + 0.11 * ((NewColor \ 65536) Mod 256)
            NewColor = RGB(NewColor, NewColor, NewColor)
        End Select

        SetPixelV hDC_Temp, MyCounterX, MyCounterY, NewColor
      End If

    Next 'MyCounterY
  Next 'MyCounterX

  ' Display the picture on the specified hDC
  BitBlt Dest_hDC, Dest_X, Dest_Y, TempBITMAP.bmWidth, TempBITMAP.bmHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy
  pvRenderBitmapGrayscale = True
CleanUp:
  ReleaseDC 0, hScreen: hScreen = 0
  SelectObject hDC_Temp, hBMP_Prev
  DeleteDC hDC_Temp: hDC_Temp = 0

  If DeletePic = True Then
    DeleteObject hNewPicture
    hNewPicture = 0
  End If

  Exit Function

pvRenderBitmapGrayscale_Error:
End Function

Private Function pvCreateIconFromBMP(ByVal hBMP_Mask As Long, _
                                     ByVal hBMP_Image As Long) As Long
  On Error GoTo pvCreateIconFromBMP_Error

  Dim TempICONINFO As ICONINFO

  If hBMP_Mask = 0 Or hBMP_Image = 0 Then Exit Function
  TempICONINFO.fIcon = 1
  TempICONINFO.hbmMask = hBMP_Mask
  TempICONINFO.hbmColor = hBMP_Image
  pvCreateIconFromBMP = CreateIconIndirect(TempICONINFO)
  Exit Function

pvCreateIconFromBMP_Error:
End Function

'*************************************************************
'
'   Private Auxiliar Subs
'
'*************************************************************
'draw a Line Using API call's
Private Sub APILine(X1 As Long, _
                    Y1 As Long, _
                    X2 As Long, _
                    Y2 As Long, _
                    lcolor As Long)
  'Use the API LineTo for Fast Drawing
  On Error GoTo APILine_Error

  Dim pt As POINT
  Dim hPen As Long, hPenOld As Long
  hPen = CreatePen(0, 1, lcolor)
  hPenOld = SelectObject(UserControl.hdc, hPen)
  MoveToEx UserControl.hdc, X1, Y1, pt
  LineTo UserControl.hdc, X2, Y2
  SelectObject UserControl.hdc, hPenOld
  DeleteObject hPen
  Exit Sub

APILine_Error:
End Sub

' full version of APILine
Private Sub APILineEx(lhdcEx As Long, _
                      X1 As Long, _
                      Y1 As Long, _
                      X2 As Long, _
                      Y2 As Long, _
                      lcolor As Long)
  'Use the API LineTo for Fast Drawing
  On Error GoTo APILineEx_Error

  Dim pt As POINT
  Dim hPen As Long, hPenOld As Long
  hPen = CreatePen(0, 1, lcolor)
  hPenOld = SelectObject(lhdcEx, hPen)
  MoveToEx lhdcEx, X1, Y1, pt
  LineTo lhdcEx, X2, Y2
  SelectObject lhdcEx, hPenOld
  DeleteObject hPen
  Exit Sub

APILineEx_Error:
End Sub

Private Sub APIFillRect(hdc As Long, _
                        RC As RECT, _
                        Color As Long)
  On Error GoTo APIFillRect_Error

  Dim NewBrush As Long
  NewBrush& = CreateSolidBrush(Color&)
  Call FillRect(hdc&, RC, NewBrush&)
  Call DeleteObject(NewBrush&)
  Exit Sub

APIFillRect_Error:
End Sub

Private Sub APIFillRectByCoords(hdc As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal w As Long, _
                                ByVal H As Long, _
                                Color As Long)
  On Error GoTo APIFillRectByCoords_Error

  Dim NewBrush As Long
  Dim tmpRect As RECT
  NewBrush& = CreateSolidBrush(Color&)
  SetRect tmpRect, X, Y, X + w, Y + H
  Call FillRect(hdc&, tmpRect, NewBrush&)
  Call DeleteObject(NewBrush&)
  Exit Sub

APIFillRectByCoords_Error:
End Sub

Private Function APIRectangle(ByVal hdc As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal w As Long, _
                              ByVal H As Long, _
                              Optional lcolor As OLE_COLOR = -1) As Long
  On Error GoTo APIRectangle_Error

  Dim hPen As Long, hPenOld As Long
  Dim pt As POINT
  hPen = CreatePen(0, 1, lcolor)
  hPenOld = SelectObject(hdc, hPen)
  MoveToEx hdc, X, Y, pt
  LineTo hdc, X + w, Y
  LineTo hdc, X + w, Y + H
  LineTo hdc, X, Y + H
  LineTo hdc, X, Y
  SelectObject hdc, hPenOld
  DeleteObject hPen
  Exit Function

APIRectangle_Error:
End Function

'Private Sub DrawCtlEdgeByRect(hdc As Long, _
'                              rt As RECT, _
'                              Optional Style As Long = EDGE_RAISED, _
'                              Optional Flags As Long = BF_RECT)
'  On Error GoTo DrawCtlEdgeByRect_Error
'
'  DrawEdge hdc, rt, Style, Flags
'  Exit Sub
'
'DrawCtlEdgeByRect_Error:
'End Sub

Private Sub DrawCtlEdge(hdc As Long, _
                        ByVal X As Single, _
                        ByVal Y As Single, _
                        ByVal w As Single, _
                        ByVal H As Single, _
                        Optional Style As Long = EDGE_RAISED, _
                        Optional ByVal Flags As Long = BF_RECT)
  On Error GoTo DrawCtlEdge_Error

  Dim R As RECT

  With R
    .Left = X
    .Top = Y
    .Right = X + w
    .Bottom = Y + H
  End With

  DrawEdge hdc, R, Style, Flags
  Exit Sub

DrawCtlEdge_Error:
End Sub

'Blend two colors
Private Function BlendColors(ByVal lcolor1 As Long, _
                             ByVal lcolor2 As Long)
  On Error GoTo BlendColors_Error

  BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)
  Exit Function

BlendColors_Error:
End Function

'System color code to long rgb
Private Function TranslateColor(ByVal lcolor As Long) As Long
  On Error GoTo TranslateColor_Error

  If OleTranslateColor(lcolor, 0, TranslateColor) Then
    TranslateColor = -1
  End If

  Exit Function

TranslateColor_Error:
End Function


Private Function MSOXPShiftColor(ByVal theColor As Long, _
                                 Optional ByVal Base As Long = &HB0) As Long
  On Error GoTo MSOXPShiftColor_Error

  Dim Red As Long, Blue As Long, Green As Long
  Dim Delta As Long
  Blue = ((theColor \ &H10000) Mod &H100)
  Green = ((theColor \ &H100) Mod &H100)
  Red = (theColor And &HFF)
  Delta = &HFF - Base
  Blue = Base + Blue * Delta \ &HFF
  Green = Base + Green * Delta \ &HFF
  Red = Base + Red * Delta \ &HFF

  If Red > 255 Then Red = 255
  If Green > 255 Then Green = 255
  If Blue > 255 Then Blue = 255
  MSOXPShiftColor = Red + 256& * Green + 65536 * Blue
  Exit Function

MSOXPShiftColor_Error:
End Function


'Offset a color
Private Function OffsetColor(lcolor As OLE_COLOR, _
                             lOffset As Long) As OLE_COLOR
  On Error GoTo OffsetColor_Error

  Dim lRed As OLE_COLOR
  Dim lGreen As OLE_COLOR
  Dim lBlue As OLE_COLOR
  Dim lR As OLE_COLOR, lg As OLE_COLOR, lb As OLE_COLOR
  lR = (lcolor And &HFF)
  lg = ((lcolor And 65280) \ 256)
  lb = ((lcolor) And 16711680) \ 65536
  lRed = (lOffset + lR)
  lGreen = (lOffset + lg)
  lBlue = (lOffset + lb)

  If lRed > 255 Then lRed = 255
  If lRed < 0 Then lRed = 0
  If lGreen > 255 Then lGreen = 255
  If lGreen < 0 Then lGreen = 0
  If lBlue > 255 Then lBlue = 255
  If lBlue < 0 Then lBlue = 0
  OffsetColor = RGB(lRed, lGreen, lBlue)
  Exit Function

OffsetColor_Error:
End Function

Private Sub DrawCaption()
  On Error GoTo DrawCaption_Error

  Dim lcolor As Long, ltmpColor As Long

  If Not m_bUseCustomColors Then
    If m_iState <> statedisabled Then
      lcolor = GetSysColor(COLOR_BTNTEXT)
    Else
      lcolor = TranslateColor(vbGrayText)
    End If

  Else

    Select Case m_iState
      Case statenormal
        lcolor = m_lFontColor
      Case statedisabled
        lcolor = TranslateColor(vbGrayText)
      Case Else
        lcolor = m_lFontHighlightColor
    End Select

  End If

  ltmpColor = UserControl.ForeColor
  UserControl.ForeColor = lcolor
  DrawText UserControl.hdc, m_sCaption, -1, m_txtRect, lwFontAlign
  UserControl.ForeColor = ltmpColor
  Exit Sub

DrawCaption_Error:
End Sub


Private Sub DrawVGradient(lEndColor As Long, _
                          lStartColor As Long, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal X2 As Long, _
                          ByVal Y2 As Long)
  ''Draw a Vertical Gradient in the current HDC
  On Error GoTo DrawVGradient_Error

  Dim dR As Single, dG As Single, dB As Single
  Dim sR As Single, sG As Single, sB As Single
  Dim eR As Single, eG As Single, eB As Single
  Dim ni As Long
  'lh = UserControl.ScaleHeight
  'lw = UserControl.ScaleWidth
  sR = (lStartColor And &HFF)
  sG = (lStartColor \ &H100) And &HFF
  sB = (lStartColor And &HFF0000) / &H10000
  eR = (lEndColor And &HFF)
  eG = (lEndColor \ &H100) And &HFF
  eB = (lEndColor And &HFF0000) / &H10000
  dR = (sR - eR) / Y2
  dG = (sG - eG) / Y2
  dB = (sB - eB) / Y2

  For ni = 0 To Y2
    APILine X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
  Next 'ni

  Exit Sub

DrawVGradient_Error:
End Sub

Private Sub DrawVGradientEx(lhdcEx As Long, _
                            lEndColor As Long, _
                            lStartColor As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long)
  ''Draw a Vertical Gradient in the current HDC
  On Error GoTo DrawVGradientEx_Error

  Dim dR As Single, dG As Single, dB As Single
  Dim sR As Single, sG As Single, sB As Single
  Dim eR As Single, eG As Single, eB As Single
  Dim ni As Long
  'lh = UserControl.ScaleHeight
  'lw = UserControl.ScaleWidth
  sR = (lStartColor And &HFF)
  sG = (lStartColor \ &H100) And &HFF
  sB = (lStartColor And &HFF0000) / &H10000
  eR = (lEndColor And &HFF)
  eG = (lEndColor \ &H100) And &HFF
  eB = (lEndColor And &HFF0000) / &H10000
  dR = (sR - eR) / Y2
  dG = (sG - eG) / Y2
  dB = (sB - eB) / Y2

  For ni = 0 To Y2
    APILineEx lhdcEx, X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
  Next 'ni

  Exit Sub

DrawVGradientEx_Error:
End Sub

Private Sub DrawHGradient(lEndColor As Long, _
                          lStartColor As Long, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal X2 As Long, _
                          ByVal Y2 As Long)
  ''Draw a Horizontal Gradient in the current HDC
  On Error GoTo DrawHGradient_Error

  Dim dR As Single, dG As Single, dB As Single
  Dim sR As Single, sG As Single, sB As Single
  Dim eR As Single, eG As Single, eB As Single
  Dim lh As Long, lw As Long
  Dim ni As Long
  lh = Y2 - Y
  lw = X2 - X
  sR = (lStartColor And &HFF)
  sG = (lStartColor \ &H100) And &HFF
  sB = (lStartColor And &HFF0000) / &H10000
  eR = (lEndColor And &HFF)
  eG = (lEndColor \ &H100) And &HFF
  eB = (lEndColor And &HFF0000) / &H10000
  dR = (sR - eR) / lw
  dG = (sG - eG) / lw
  dB = (sB - eB) / lw

  For ni = 0 To lw
    APILine X + ni, Y, X + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
  Next 'ni

  Exit Sub

DrawHGradient_Error:
End Sub

Private Sub DrawJavaBorder(ByVal X As Long, _
                           ByVal Y As Long, _
                           ByVal w As Long, _
                           ByVal H As Long, _
                           ByVal lColorShadow As Long, _
                           ByVal lColorLight As Long, _
                           ByVal lColorBack As Long)
  On Error GoTo DrawJavaBorder_Error

  APIRectangle UserControl.hdc, X, Y, w - 1, H - 1, lColorShadow
  APIRectangle UserControl.hdc, X + 1, Y + 1, w - 1, H - 1, lColorLight
  SetPixelV UserControl.hdc, X, Y + H, lColorBack
  SetPixelV UserControl.hdc, X + w, Y, lColorBack
  SetPixelV UserControl.hdc, X + 1, Y + H - 1, BlendColors(lColorLight, lColorShadow)
  SetPixelV UserControl.hdc, X + w - 1, Y + 1, BlendColors(lColorLight, lColorShadow)
  Exit Sub

DrawJavaBorder_Error:
End Sub
    
Private Function DrawTheme(sClass As String, _
                           ByVal iPart As Long, _
                           ByVal iState As Long) As Boolean
  
  Dim hTheme As Long
  Dim lResult As Long
  Dim m_btnRect2 As RECT
  Dim hRgn As Long
  On Error GoTo NoXP

  hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))

    If hTheme Then
        If m_bRoundedBordersByTheme Then
            '<--Rounded Region as requested for some themes:
            'Thanks to Dana Seaman-->
            SetRect m_btnRect2, m_btnRect.Left - 1, m_btnRect.Top - 1, m_btnRect.Right + 1, m_btnRect.Bottom + 1
            lResult = GetThemeBackgroundRegion(hTheme, UserControl.hdc, iPart, iState, m_btnRect2, hRgn)
            SetWindowRgn hWnd, hRgn, True
            'free the memory.
            DeleteObject hRgn
        End If
        lResult = DrawThemeBackground(hTheme, UserControl.hdc, iPart, iState, m_btnRect, m_btnRect)
        DrawTheme = IIf(lResult, False, True)
    Else
        DrawTheme = False
    End If

  Exit Function

NoXP:
  DrawTheme = False

End Function

Private Function CreateWinXPregion() As Long
  On Error GoTo CreateWinXPRegion_Error

  Dim pPoligon(8) As POINT
  Dim cpPoligon(1) As Long
  Dim lw As Long, lh As Long
  lw = UserControl.ScaleWidth
  lh = UserControl.ScaleHeight
  cpPoligon(0) = 5
  cpPoligon(1) = 5
  pPoligon(0).X = 0: pPoligon(0).Y = 1
  pPoligon(1).X = 1: pPoligon(1).Y = 0
  pPoligon(2).X = lw - 1: pPoligon(2).Y = 0
  pPoligon(3).X = lw: pPoligon(3).Y = 1
  pPoligon(4).X = lw: pPoligon(4).Y = lh - 2
  pPoligon(5).X = lw - 2: pPoligon(5).Y = lh
  pPoligon(6).X = 2: pPoligon(6).Y = lh
  pPoligon(7).X = 0: pPoligon(7).Y = lh - 2
  'pPoligon(8).x = 0: pPoligon(8).y = lh - 2
  CreateWinXPregion = CreatePolygonRgn(pPoligon(0), 8, ALTERNATE)
  Exit Function

CreateWinXPRegion_Error:
End Function

Private Function CreateGalaxyRegion() As Long
  On Error GoTo CreateGalaxyRegion_Error

  Dim pPoligon(8) As POINT
  Dim cpPoligon(1) As Long
  Dim lw As Long, lh As Long
  lw = UserControl.ScaleWidth
  lh = UserControl.ScaleHeight
  cpPoligon(0) = 5
  cpPoligon(1) = 5
  pPoligon(0).X = 0: pPoligon(0).Y = 2
  pPoligon(1).X = 2: pPoligon(1).Y = 0
  pPoligon(2).X = lw - 3: pPoligon(2).Y = 0
  pPoligon(3).X = lw: pPoligon(3).Y = 3
  pPoligon(4).X = lw: pPoligon(4).Y = lh - 3
  pPoligon(5).X = lw - 3: pPoligon(5).Y = lh
  pPoligon(6).X = 4: pPoligon(6).Y = lh
  pPoligon(7).X = 0: pPoligon(7).Y = lh - 4
  'pPoligon(8).x = 0: pPoligon(8).y = lh - 2
  CreateGalaxyRegion = CreatePolygonRgn(pPoligon(0), 8, ALTERNATE)
  Exit Function

CreateGalaxyRegion_Error:
End Function

Private Function CreateMacOSXButtonRegion() As Long
  'MsgBox "MACOS?"
  On Error GoTo CreateMacOSXButtonRegion_Error

  CreateMacOSXButtonRegion = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1, 18, 18)
  Exit Function

CreateMacOSXButtonRegion_Error:
End Function

Public Sub About()
  On Error GoTo About_Error

  m_About.Visible = True
  SetWindowLongA m_About.hWnd, GWL_STYLE, lPrevStyle + WS_CAPTION + WS_THICKFRAME + WS_MINIMIZEBOX
  SetWindowPos m_About.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW 'Or SWP_NOACTIVATE
  SetWindowPos m_About.hWnd, 0, 0, 0, 0, 0, SWP_REFRESH
  SetWindowText m_About.hWnd, "About isButton " & strCurrentVersion
  SetWindowPos m_About.hWnd, 0, 0, 0, 0, 0, SWP_REFRESH
  SetParent m_About.hWnd, 0
  ''This is the ocx version about dialog
  'frmAbout.Show vbModal
  Exit Sub

About_Error:
End Sub

'****************************************************************
'
'   Procedures
'
'****************************************************************
Private Sub DrawWinXPButton(Mode As isState)
  '' This Sub Draws the XPStyle Button
  On Error GoTo DrawWinXPButton_Error

  Dim lHDC As Long
  Dim tempColor As Long
  Dim lh As Long, lw As Long
  Dim lcw As Long, lch As Long
  Dim lStep As Single
  lw = UserControl.ScaleWidth
  lh = UserControl.ScaleHeight
  lHDC = UserControl.hdc
  lcw = m_btnRect.Left + lw / 2 + 1
  lch = m_btnRect.Top + lh / 2
  lStep = 25 / lh
  UserControl.BackColor = GetSysColor(COLOR_BTNFACE)

  Select Case Mode
    Case statenormal, stateHot:
      'Main
      DrawVGradient &HFBFCFC, &HF0F0F0, 1, 1, lw - 2, 4
      DrawVGradient &HF9FAFA, &HEAF0F0, 1, 4, lw - 2, lh - 8
      DrawVGradient &HE6EBEB, &HC5D0D6, 1, lh - 4, lw - 2, 3
      'right
      DrawVGradient &HFAFBFB, &HDAE2E4, lw - 3, 3, lw - 2, lh - 5
      DrawVGradient &HF2F4F5, &HCDD7DB, lw - 2, 3, lw - 1, lh - 5
      'Border
      APILine 1, 0, lw - 1, 0, &H743C00
      APILine 0, 1, 0, lh - 1, &H743C00
      APILine lw - 1, 1, lw - 1, lh - 1, &H743C00
      APILine 1, lh - 1, lw - 1, lh - 1, &H743C00
      'Corners
      SetPixelV lHDC, 1, 1, &H906E48
      SetPixelV lHDC, 1, lh - 2, &H906E48
      SetPixelV lHDC, lw - 2, 1, &H906E48
      SetPixelV lHDC, lw - 2, lh - 2, &H906E48
      'External Borders
      SetPixelV lHDC, 0, 1, &HA28B6A
      SetPixelV lHDC, 1, 0, &HA28B6A
      SetPixelV lHDC, 1, lh - 1, &HA28B6A
      SetPixelV lHDC, 0, lh - 2, &HA28B6A
      SetPixelV lHDC, lw - 1, lh - 2, &HA28B6A
      SetPixelV lHDC, lw - 2, lh - 1, &HA28B6A
      SetPixelV lHDC, lw - 2, 0, &HA28B6A
      SetPixelV lHDC, lw - 1, 1, &HA28B6A
      'Internal Soft
      SetPixelV lHDC, 1, 2, &HCAC7BF
      SetPixelV lHDC, 2, 1, &HCAC7BF
      SetPixelV lHDC, 2, lh - 2, &HCAC7BF
      SetPixelV lHDC, 1, lh - 3, &HCAC7BF
      SetPixelV lHDC, lw - 2, lh - 3, &HCAC7BF
      SetPixelV lHDC, lw - 3, lh - 2, &HCAC7BF
      SetPixelV lHDC, lw - 3, 1, &HCAC7BF
      SetPixelV lHDC, lw - 2, 2, &HCAC7BF

      If Mode = stateHot Then
        APILine 2, 1, lw - 2, 1, &HCFF0FF
        APILine 2, 2, lw - 2, 2, &H89D8FD
        APILine 2, lh - 3, lw - 2, lh - 3, &H30B3F8
        APILine 2, lh - 2, lw - 2, lh - 2, &H1097E5
        DrawVGradient &H89D8FD, &H30B3F8, 1, 2, 3, lh - 5
        DrawVGradient &H89D8FD, &H30B3F8, lw - 3, 2, lw - 1, lh - 5
      ElseIf (Mode = statenormal And m_bFocused) Or Ambient.DisplayAsDefault Then
        APILine 2, lh - 2, lw - 2, lh - 2, &HEE8269
        APILine 2, 1, lw - 2, 1, &HFFE7CE
        APILine 2, 2, lw - 2, 2, &HF6D4BC
        APILine 2, lh - 3, lw - 2, lh - 3, &HE4AD89
        DrawVGradient &HF6D4BC, &HE4AD89, 1, 2, 3, lh - 5
        DrawVGradient &HF6D4BC, &HE4AD89, lw - 3, 2, lw - 1, lh - 5
      End If

    Case statePressed:
      ' &HC1ccD1 - &HDBE2E3   -&HDCE3E4   -&HC1CCD1   -&HEEF1F2
      'Main
      DrawVGradient &HC1CCD1, &HDCE3E4, 2, 1, lw - 1, 4
      DrawVGradient &HDCE3E4, &HDBE2E3, 2, 4, lw - 1, lh - 8
      DrawVGradient &HDBE2E3, &HEEF1F2, 3, lh - 4, lw - 1, 3
      'left
      DrawVGradient &HCED8DA, &HDBE2E3, 1, 3, 2, lh - 5
      DrawVGradient &HCED8DA, &HDBE2E3, 2, 4, 3, lh - 7
      'Border
      APILine 1, 0, lw - 1, 0, &H743C00
      APILine 0, 1, 0, lh - 1, &H743C00
      APILine lw - 1, 1, lw - 1, lh - 1, &H743C00
      APILine 1, lh - 1, lw - 1, lh - 1, &H743C00
      'Corners
      SetPixelV lHDC, 1, 1, &H906E48
      SetPixelV lHDC, 1, lh - 2, &H906E48
      SetPixelV lHDC, lw - 2, 1, &H906E48
      SetPixelV lHDC, lw - 2, lh - 2, &H906E48
      'External Borders
      SetPixelV lHDC, 0, 1, &HA28B6A
      SetPixelV lHDC, 1, 0, &HA28B6A
      SetPixelV lHDC, 1, lh - 1, &HA28B6A
      SetPixelV lHDC, 0, lh - 2, &HA28B6A
      SetPixelV lHDC, lw - 1, lh - 2, &HA28B6A
      SetPixelV lHDC, lw - 2, lh - 1, &HA28B6A
      SetPixelV lHDC, lw - 2, 0, &HA28B6A
      SetPixelV lHDC, lw - 1, 1, &HA28B6A
      'Internal Soft
      SetPixelV lHDC, 1, 2, &HCAC7BF
      SetPixelV lHDC, 2, 1, &HCAC7BF
      SetPixelV lHDC, 2, lh - 2, &HCAC7BF
      SetPixelV lHDC, 1, lh - 3, &HCAC7BF
      SetPixelV lHDC, lw - 2, lh - 3, &HCAC7BF
      SetPixelV lHDC, lw - 3, lh - 2, &HCAC7BF
      SetPixelV lHDC, lw - 3, 1, &HCAC7BF
      SetPixelV lHDC, lw - 2, 2, &HCAC7BF
    Case statedisabled:
      tempColor = &HEAF4F5
      UserControl.BackColor = tempColor
      lHDC = UserControl.hdc
      APIRectangle lHDC, 0, 0, lw - 1, lh - 1, &HBAC7C9
      tempColor = &HC7D5D8
      SetPixelV lHDC, 0, 1, tempColor
      SetPixelV lHDC, 1, 1, tempColor
      SetPixelV lHDC, 1, 0, tempColor
      SetPixelV lHDC, 0, lh - 2, tempColor
      SetPixelV lHDC, 1, lh - 2, tempColor
      SetPixelV lHDC, 1, lh - 1, tempColor
      SetPixelV lHDC, lw - 1, 1, tempColor
      SetPixelV lHDC, lw - 2, 1, tempColor
      SetPixelV lHDC, lw - 2, 0, tempColor
      SetPixelV lHDC, lw - 1, lh - 2, tempColor
      SetPixelV lHDC, lw - 2, lh - 2, tempColor
      SetPixelV lHDC, lw - 2, lh - 1, tempColor
  End Select

  Exit Sub

DrawWinXPButton_Error:
End Sub

Private Sub DrawCustomWinXPButton(Mode As isState)
  On Error GoTo DrawCustomWinXPButton_Error

  Dim tmpcolor As Long
  Dim lh As Long, lw As Long
  Dim lHDC As Long
  lh = UserControl.ScaleHeight: lw = UserControl.ScaleWidth

  'Here, we know we will use custom colors
  Select Case Mode
    Case statenormal, stateDefaulted, stateHot
      tmpcolor = m_lBackColor
      UserControl.BackColor = tmpcolor
      'main gradient
      DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&HF), 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3
      DrawVGradient OffsetColor(tmpcolor, &H15), OffsetColor(tmpcolor, -&H15), 1, 2, 2, UserControl.ScaleHeight - 5
      DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H20), UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5
      'Top Lines
      APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(tmpcolor, &H5)
      APILine 1, 2, UserControl.ScaleWidth - 1, 2, OffsetColor(tmpcolor, &H2)
      'Bottom Lines
      APILine 2, UserControl.ScaleHeight - 4, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 4, OffsetColor(tmpcolor, -&H10)
      APILine 2, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H18)
      APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(tmpcolor, -&H25)
      'Border
      tmpcolor = OffsetColor(tmpcolor, -&H80)
      APILine 2, 0, UserControl.ScaleWidth - 2, 0, tmpcolor
      APILine 0, 2, 0, UserControl.ScaleHeight - 2, tmpcolor
      APILine 2, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      APILine UserControl.ScaleWidth - 1, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight, tmpcolor
      SetPixelV UserControl.hdc, 1, 1, tmpcolor: SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 2, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, tmpcolor
      'Border Pixels
      tmpcolor = OffsetColor(m_lBackColor, -&H15)
      SetPixelV UserControl.hdc, 1, 0, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 0, tmpcolor
      SetPixelV UserControl.hdc, 0, 1, tmpcolor: SetPixelV UserControl.hdc, 0, UserControl.ScaleHeight - 2, tmpcolor
      SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 2, tmpcolor

      If Mode = stateDefaulted Or Mode = stateHot Or (m_bFocused And m_bShowFocus) Then 'Or Ambient.DisplayAsDefault  Then
        tmpcolor = IIf((Mode = stateHot), m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
        APILine 2, 1, lw - 2, 1, OffsetColor(tmpcolor, &H55)
        APILine 2, 2, lw - 2, 2, OffsetColor(tmpcolor, &H45)
        APILine 2, lh - 3, lw - 2, lh - 3, OffsetColor(tmpcolor, &H10)
        APILine 2, lh - 2, lw - 2, lh - 2, tmpcolor
        DrawVGradient OffsetColor(tmpcolor, &H45), OffsetColor(tmpcolor, &H10), 1, 2, 3, lh - 5
        DrawVGradient OffsetColor(tmpcolor, &H45), OffsetColor(tmpcolor, &H10), lw - 3, 2, lw - 1, lh - 5
      End If

    Case statePressed:
      tmpcolor = m_lBackColor
      lHDC = UserControl.hdc
      'Main
      DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H15), 2, 1, lw - 1, 4
      DrawVGradient OffsetColor(tmpcolor, -&H15), OffsetColor(tmpcolor, -&H5), 2, 4, lw - 1, lh - 8
      DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, &H5), 3, lh - 4, lw - 1, 3
      'left
      DrawVGradient OffsetColor(tmpcolor, -&H20), OffsetColor(tmpcolor, -&H16), 1, 3, 2, lh - 5
      DrawVGradient OffsetColor(tmpcolor, -&H18), OffsetColor(tmpcolor, -&HF), 2, 4, 3, lh - 7
      'External Borders
      'tmpcolor = vbBlue
      SetPixelV lHDC, 1, 2, OffsetColor(tmpcolor, -&H30)
      SetPixelV lHDC, 2, 1, OffsetColor(tmpcolor, -&H30)
      SetPixelV lHDC, 2, lh - 2, OffsetColor(tmpcolor, -&H5)
      SetPixelV lHDC, 1, lh - 3, OffsetColor(tmpcolor, -&H10)
      SetPixelV lHDC, lw - 2, lh - 3, OffsetColor(tmpcolor, &H12)
      SetPixelV lHDC, lw - 3, lh - 2, OffsetColor(tmpcolor, &H8)
      SetPixelV lHDC, lw - 3, 1, OffsetColor(tmpcolor, -&H30)
      SetPixelV lHDC, lw - 2, 2, OffsetColor(tmpcolor, -&H25)
      'Border
      tmpcolor = OffsetColor(m_lBackColor, -&H80)
      APILine 1, 0, lw - 1, 0, tmpcolor
      APILine 0, 1, 0, lh - 1, tmpcolor
      APILine lw - 1, 1, lw - 1, lh - 1, tmpcolor
      APILine 1, lh - 1, lw - 1, lh - 1, tmpcolor
      'Corners
      tmpcolor = OffsetColor(m_lBackColor, -&H60)
      SetPixelV lHDC, 1, 1, tmpcolor
      SetPixelV lHDC, 1, lh - 2, tmpcolor
      SetPixelV lHDC, lw - 2, 1, tmpcolor
      SetPixelV lHDC, lw - 2, lh - 2, tmpcolor
    Case statedisabled
      tmpcolor = m_lBackColor
      UserControl.BackColor = tmpcolor
      lHDC = UserControl.hdc
      APIRectangle lHDC, 0, 0, lw - 1, lh - 1, OffsetColor(m_lBackColor, -&H40)
      tmpcolor = OffsetColor(m_lBackColor, -&H35)
      SetPixelV lHDC, 0, 1, tmpcolor
      SetPixelV lHDC, 1, 1, tmpcolor
      SetPixelV lHDC, 1, 0, tmpcolor
      SetPixelV lHDC, 0, lh - 2, tmpcolor
      SetPixelV lHDC, 1, lh - 2, tmpcolor
      SetPixelV lHDC, 1, lh - 1, tmpcolor
      SetPixelV lHDC, lw - 1, 1, tmpcolor
      SetPixelV lHDC, lw - 2, 1, tmpcolor
      SetPixelV lHDC, lw - 2, 0, tmpcolor
      SetPixelV lHDC, lw - 1, lh - 2, tmpcolor
      SetPixelV lHDC, lw - 2, lh - 2, tmpcolor
      SetPixelV lHDC, lw - 2, lh - 1, tmpcolor
  End Select

  Exit Sub

DrawCustomWinXPButton_Error:
End Sub

Private Sub DrawMacOSXButton()
  On Error GoTo DrawMacOSXButton_Error

  If m_iState = stateHot Or m_iState = stateDefaulted Then ' Or Ambient.DisplayAsDefault Then
    DrawMacOSXButtonHot
  ElseIf m_iState = statenormal Or m_iState = statedisabled Then

    If m_bFocused Then 'Or Ambient.DisplayAsDefault Then
      DrawMacOSXButtonHot
    Else
      DrawMacOSXButtonNormal
    End If

  Else 'If m_iState = statePressed Then
    DrawMacOSXButtonPressed
  End If

  Exit Sub

DrawMacOSXButton_Error:
End Sub

Private Sub DrawMacOSXButtonNormal()
  On Error GoTo DrawMacOSXButtonNormal_Error

  Dim lHDC As Long
  lHDC = UserControl.hdc
  'Variable vars (real into code)
  Dim lh As Long, lw As Long
  lh = UserControl.ScaleHeight: lw = UserControl.ScaleWidth
  Dim tmph As Long, tmpw As Long
  Dim tmph1 As Long, tmpw1 As Long
  APIFillRectByCoords hdc, 18, 11, lw - 34, lh - 19, &HEAE7E8
  SetPixelV lHDC, 6, 0, &HFEFEFE: SetPixelV lHDC, 7, 0, &HE6E6E6: SetPixelV lHDC, 8, 0, &HACACAC: SetPixelV lHDC, 9, 0, &H7A7A7A: SetPixelV lHDC, 10, 0, &H6C6C6C: SetPixelV lHDC, 11, 0, &H6B6B6B: SetPixelV lHDC, 12, 0, &H6F6F6F: SetPixelV lHDC, 13, 0, &H716F6F: SetPixelV lHDC, 14, 0, &H727070: SetPixelV lHDC, 15, 0, &H676866: SetPixelV lHDC, 16, 0, &H6C6D6B: SetPixelV lHDC, 17, 0, &H67696A: SetPixelV lHDC, 5, 1, &HEFEFEF: SetPixelV lHDC, 6, 1, &H939393: SetPixelV lHDC, 7, 1, &H676767: SetPixelV lHDC, 8, 1, &H797979: SetPixelV lHDC, 9, 1, &HB3B3B3: SetPixelV lHDC, 10, 1, &HDBDBDB: SetPixelV lHDC, 11, 1, &HEBEDEE: SetPixelV lHDC, 12, 1, &HF5F4F6: SetPixelV lHDC, 13, 1, &HF5F4F6: SetPixelV lHDC, 14, 1, &HF5F4F6: SetPixelV lHDC, 15, 1, &HF5F4F6: SetPixelV lHDC, 16, 1, &HF5F4F6: SetPixelV lHDC, 17, 1, &HF5F4F6
  SetPixelV lHDC, 3, 2, &HFEFEFE: SetPixelV lHDC, 4, 2, &HE5E5E5: SetPixelV lHDC, 5, 2, &H737373: SetPixelV lHDC, 6, 2, &H656565: SetPixelV lHDC, 7, 2, &H939393: SetPixelV lHDC, 8, 2, &HDCDCDC: SetPixelV lHDC, 9, 2, &HE9E9E9: SetPixelV lHDC, 10, 2, &HF2F1F3: SetPixelV lHDC, 11, 2, &HF3F2F4: SetPixelV lHDC, 12, 2, &HF2F1F3: SetPixelV lHDC, 13, 2, &HF3F2F4: SetPixelV lHDC, 14, 2, &HF2F1F3: SetPixelV lHDC, 15, 2, &HF3F2F4: SetPixelV lHDC, 16, 2, &HF2F1F3: SetPixelV lHDC, 17, 2, &HF3F2F4: SetPixelV lHDC, 3, 3, &HEEEEEE: SetPixelV lHDC, 4, 3, &H717171: SetPixelV lHDC, 5, 3, &H6C6C6C: SetPixelV lHDC, 6, 3, &H909090: SetPixelV lHDC, 7, 3, &HD2D2D2: SetPixelV lHDC, 8, 3, &HE3E3E3: SetPixelV lHDC, 9, 3, &HECECEC: SetPixelV lHDC, 10, 3, &HEDEDED: SetPixelV lHDC, 11, 3, &HEEEEEE: SetPixelV lHDC, 12, 3, &HEDEDED: SetPixelV lHDC, 13, 3, &HEEEEEE: SetPixelV lHDC, 14, 3, &HEDEDED: SetPixelV lHDC, 15, 3, &HEEEEEE: SetPixelV lHDC, 16, 3, &HEDEDED: SetPixelV lHDC, 17, 3, &HEEEEEE
  SetPixelV lHDC, 2, 4, &HFBFBFB: SetPixelV lHDC, 3, 4, &H858585: SetPixelV lHDC, 4, 4, &H686868: SetPixelV lHDC, 5, 4, &H959595: SetPixelV lHDC, 6, 4, &HB1B1B1: SetPixelV lHDC, 7, 4, &HDCDCDC: SetPixelV lHDC, 8, 4, &HE3E3E3: SetPixelV lHDC, 9, 4, &HE3E3E3: SetPixelV lHDC, 10, 4, &HEAEAEA: SetPixelV lHDC, 11, 4, &HEBEBEB: SetPixelV lHDC, 12, 4, &HEBEBEB: SetPixelV lHDC, 13, 4, &HEBEBEB: SetPixelV lHDC, 14, 4, &HEBEBEB: SetPixelV lHDC, 15, 4, &HEBEBEB: SetPixelV lHDC, 16, 4, &HEBEBEB: SetPixelV lHDC, 17, 4, &HEBEBEB:
  SetPixelV lHDC, 1, 5, &HFEFEFE: SetPixelV lHDC, 2, 5, &HCACACA: SetPixelV lHDC, 3, 5, &H696969: SetPixelV lHDC, 4, 5, &H949494: SetPixelV lHDC, 5, 5, &HA6A6A6: SetPixelV lHDC, 6, 5, &HC5C5C5: SetPixelV lHDC, 7, 5, &HD8D8D8: SetPixelV lHDC, 8, 5, &HE0E0E0: SetPixelV lHDC, 9, 5, &HE1E1E1: SetPixelV lHDC, 10, 5, &HEAE9EA: SetPixelV lHDC, 11, 5, &HE7E7E7: SetPixelV lHDC, 12, 5, &HE9E7E8: SetPixelV lHDC, 13, 5, &HEBE8EA: SetPixelV lHDC, 14, 5, &HEAE7E9: SetPixelV lHDC, 15, 5, &HEBE8EA: SetPixelV lHDC, 16, 5, &HEAE7E9: SetPixelV lHDC, 17, 5, &HEBE8EA
  SetPixelV lHDC, 1, 6, &HF9F9F9: SetPixelV lHDC, 2, 6, &H808080: SetPixelV lHDC, 3, 6, &H878787: SetPixelV lHDC, 4, 6, &HA8A8A8: SetPixelV lHDC, 5, 6, &HB3B3B3: SetPixelV lHDC, 6, 6, &HC6C6C6: SetPixelV lHDC, 7, 6, &HDEDEDE: SetPixelV lHDC, 8, 6, &HE0E0E0: SetPixelV lHDC, 9, 6, &HE2E2E2: SetPixelV lHDC, 10, 6, &HE3E2E2: SetPixelV lHDC, 11, 6, &HE9EAE9: SetPixelV lHDC, 12, 6, &HE9E8E9: SetPixelV lHDC, 13, 6, &HEBE8EA: SetPixelV lHDC, 14, 6, &HEBE8EA: SetPixelV lHDC, 15, 6, &HEBE8EA: SetPixelV lHDC, 16, 6, &HEBE8EA: SetPixelV lHDC, 17, 6, &HEBE8EA
  SetPixelV lHDC, 1, 7, &HE8E8E8: SetPixelV lHDC, 2, 7, &H777777: SetPixelV lHDC, 3, 7, &H9B9B9B: SetPixelV lHDC, 4, 7, &HB1B1B1: SetPixelV lHDC, 5, 7, &HB9B9B9: SetPixelV lHDC, 6, 7, &HC5C5C5: SetPixelV lHDC, 7, 7, &HD6D6D6: SetPixelV lHDC, 8, 7, &HE0E0E0: SetPixelV lHDC, 9, 7, &HE0E0E0: SetPixelV lHDC, 10, 7, &HE7E7E7: SetPixelV lHDC, 11, 7, &HE7E7E7: SetPixelV lHDC, 12, 7, &HE9E9E9: SetPixelV lHDC, 13, 7, &HEAEAEA: SetPixelV lHDC, 14, 7, &HEAEAEA: SetPixelV lHDC, 15, 7, &HEAEAEA: SetPixelV lHDC, 16, 7, &HEAEAEA: SetPixelV lHDC, 17, 7, &HEAEAEA
  SetPixelV lHDC, 0, 8, &HFDFDFD: SetPixelV lHDC, 1, 8, &HC6C6C6: SetPixelV lHDC, 2, 8, &H7E7E7E: SetPixelV lHDC, 3, 8, &HABABAB: SetPixelV lHDC, 4, 8, &HC1C1C1: SetPixelV lHDC, 5, 8, &HC1C1C1: SetPixelV lHDC, 6, 8, &HCBCBCB: SetPixelV lHDC, 7, 8, &HCECECE: SetPixelV lHDC, 8, 8, &HD5D5D5: SetPixelV lHDC, 9, 8, &HD8D8D8: SetPixelV lHDC, 10, 8, &HDADADA: SetPixelV lHDC, 11, 8, &HDDDDDD: SetPixelV lHDC, 12, 8, &HDEDEDE: SetPixelV lHDC, 13, 8, &HE1E1E1: SetPixelV lHDC, 14, 8, &HE0E0E0: SetPixelV lHDC, 15, 8, &HE1E1E1: SetPixelV lHDC, 16, 8, &HE0E0E0: SetPixelV lHDC, 17, 8, &HE1E1E1
  SetPixelV lHDC, 0, 9, &HFAFAFA: SetPixelV lHDC, 1, 9, &HAEAEAE: SetPixelV lHDC, 2, 9, &H919191: SetPixelV lHDC, 3, 9, &HB9B9B9: SetPixelV lHDC, 4, 9, &HC4C4C4: SetPixelV lHDC, 5, 9, &HCECECE: SetPixelV lHDC, 6, 9, &HD1D1D1: SetPixelV lHDC, 7, 9, &HDADADA: SetPixelV lHDC, 8, 9, &HDCDCDC: SetPixelV lHDC, 9, 9, &HDBDBDB: SetPixelV lHDC, 10, 9, &HDFDFDF: SetPixelV lHDC, 11, 9, &HE1E3E1: SetPixelV lHDC, 12, 9, &HE2E3E2: SetPixelV lHDC, 13, 9, &HE5E2E3: SetPixelV lHDC, 14, 9, &HE5E2E3: SetPixelV lHDC, 15, 9, &HE5E2E3: SetPixelV lHDC, 16, 9, &HE5E2E3: SetPixelV lHDC, 17, 9, &HE5E2E3
  SetPixelV lHDC, 0, 10, &HF7F7F7: SetPixelV lHDC, 1, 10, &HA0A0A0: SetPixelV lHDC, 2, 10, &H999999: SetPixelV lHDC, 3, 10, &HC3C3C3: SetPixelV lHDC, 4, 10, &HC9C9C9: SetPixelV lHDC, 5, 10, &HD5D5D5: SetPixelV lHDC, 6, 10, &HD7D7D7: SetPixelV lHDC, 7, 10, &HDFDFDF: SetPixelV lHDC, 8, 10, &HE0E0E0: SetPixelV lHDC, 9, 10, &HE0E0E0: SetPixelV lHDC, 10, 10, &HE4E4E4: SetPixelV lHDC, 11, 10, &HE6E8E6: SetPixelV lHDC, 12, 10, &HE8E7E7: SetPixelV lHDC, 13, 10, &HEAE7E8: SetPixelV lHDC, 14, 10, &HEAE7E8: SetPixelV lHDC, 15, 10, &HEAE7E8: SetPixelV lHDC, 16, 10, &HEAE7E8: SetPixelV lHDC, 17, 10, &HEAE7E8
  SetPixelV lHDC, 0, 11, &HF5F5F5: SetPixelV lHDC, 1, 11, &HA3A3A3: SetPixelV lHDC, 2, 11, &H9B9B9B: SetPixelV lHDC, 3, 11, &HC6C6C6: SetPixelV lHDC, 4, 11, &HD3D3D3: SetPixelV lHDC, 5, 11, &HD6D6D6: SetPixelV lHDC, 6, 11, &HDDDDDD: SetPixelV lHDC, 7, 11, &HE1E1E1: SetPixelV lHDC, 8, 11, &HE3E3E3: SetPixelV lHDC, 9, 11, &HE6E6E6: SetPixelV lHDC, 10, 11, &HE7E8E7: SetPixelV lHDC, 11, 11, &HE9EAE9: SetPixelV lHDC, 12, 11, &HE8EAE9: SetPixelV lHDC, 13, 11, &HE8EBE9: SetPixelV lHDC, 14, 11, &HE8EBE9: SetPixelV lHDC, 15, 11, &HE8EBE9: SetPixelV lHDC, 16, 11, &HE8EBE9: SetPixelV lHDC, 17, 11, &HE8EBE9
  SetPixelV lHDC, 0, 12, &HF5F5F5: SetPixelV lHDC, 1, 12, &HAAAAAA: SetPixelV lHDC, 2, 12, &H8E8E8E: SetPixelV lHDC, 3, 12, &HD0D0D0: SetPixelV lHDC, 4, 12, &HDADADA: SetPixelV lHDC, 5, 12, &HDFDFDF: SetPixelV lHDC, 6, 12, &HE4E4E4: SetPixelV lHDC, 7, 12, &HE6E6E6: SetPixelV lHDC, 8, 12, &HE8E8E8: SetPixelV lHDC, 9, 12, &HECECEC: SetPixelV lHDC, 10, 12, &HEEEFEE: SetPixelV lHDC, 11, 12, &HEEF0EF: SetPixelV lHDC, 12, 12, &HEEF0EF: SetPixelV lHDC, 13, 12, &HEEF1EF: SetPixelV lHDC, 14, 12, &HEEF1EF: SetPixelV lHDC, 15, 12, &HEEF1EF: SetPixelV lHDC, 16, 12, &HEEF1EF: SetPixelV lHDC, 17, 12, &HEEF1EF
  tmph = lh - 22
  SetPixelV lHDC, 0, tmph + 12, &HF5F5F5: SetPixelV lHDC, 1, tmph + 12, &HAAAAAA: SetPixelV lHDC, 2, tmph + 12, &H8E8E8E: SetPixelV lHDC, 3, tmph + 12, &HD0D0D0: SetPixelV lHDC, 4, tmph + 12, &HDADADA: SetPixelV lHDC, 5, tmph + 12, &HDFDFDF: SetPixelV lHDC, 6, tmph + 12, &HE4E4E4: SetPixelV lHDC, 7, tmph + 12, &HE6E6E6: SetPixelV lHDC, 8, tmph + 12, &HE8E8E8: SetPixelV lHDC, 9, tmph + 12, &HECECEC: SetPixelV lHDC, 10, tmph + 12, &HEEEFEE: SetPixelV lHDC, 11, tmph + 12, &HEEF0EF: SetPixelV lHDC, 12, tmph + 12, &HEEF0EF: SetPixelV lHDC, 13, tmph + 12, &HEEF1EF: SetPixelV lHDC, 14, tmph + 12, &HEEF1EF: SetPixelV lHDC, 15, tmph + 12, &HEEF1EF: SetPixelV lHDC, 16, tmph + 12, &HEEF1EF: SetPixelV lHDC, 17, tmph + 12, &HEEF1EF
  SetPixelV lHDC, 0, tmph + 13, &HF7F7F7: SetPixelV lHDC, 1, tmph + 13, &HC2C2C2: SetPixelV lHDC, 2, tmph + 13, &H838383: SetPixelV lHDC, 3, tmph + 13, &HCFCFCF: SetPixelV lHDC, 4, tmph + 13, &HDEDEDE: SetPixelV lHDC, 5, tmph + 13, &HE3E3E3: SetPixelV lHDC, 6, tmph + 13, &HE8E8E8: SetPixelV lHDC, 7, tmph + 13, &HEAEAEA: SetPixelV lHDC, 8, tmph + 13, &HEDEDED: SetPixelV lHDC, 9, tmph + 13, &HF1F1F1: SetPixelV lHDC, 10, tmph + 13, &HF2F2F2: SetPixelV lHDC, 11, tmph + 13, &HF2F2F2: SetPixelV lHDC, 12, tmph + 13, &HF2F2F2: SetPixelV lHDC, 13, tmph + 13, &HF2F2F2: SetPixelV lHDC, 14, tmph + 13, &HF2F2F2: SetPixelV lHDC, 15, tmph + 13, &HF2F2F2: SetPixelV lHDC, 16, tmph + 13, &HF2F2F2: SetPixelV lHDC, 17, tmph + 13, &HF2F2F2
  SetPixelV lHDC, 0, tmph + 14, &HFBFBFB: SetPixelV lHDC, 1, tmph + 14, &HE1E1E1: SetPixelV lHDC, 2, tmph + 14, &H818181: SetPixelV lHDC, 3, tmph + 14, &HABABAB: SetPixelV lHDC, 4, tmph + 14, &HDCDCDC: SetPixelV lHDC, 5, tmph + 14, &HE5E5E5: SetPixelV lHDC, 6, tmph + 14, &HEDEDED: SetPixelV lHDC, 7, tmph + 14, &HEFEFEF: SetPixelV lHDC, 8, tmph + 14, &HF1F1F1: SetPixelV lHDC, 9, tmph + 14, &HF4F4F4: SetPixelV lHDC, 10, tmph + 14, &HF5F5F5: SetPixelV lHDC, 11, tmph + 14, &HF5F5F5: SetPixelV lHDC, 12, tmph + 14, &HF5F5F5: SetPixelV lHDC, 13, tmph + 14, &HF5F5F5: SetPixelV lHDC, 14, tmph + 14, &HF5F5F5: SetPixelV lHDC, 15, tmph + 14, &HF5F5F5: SetPixelV lHDC, 16, tmph + 14, &HF5F5F5: SetPixelV lHDC, 17, tmph + 14, &HF5F5F5
  SetPixelV lHDC, 0, tmph + 15, &HFEFEFE: SetPixelV lHDC, 1, tmph + 15, &HEDEDED: SetPixelV lHDC, 2, tmph + 15, &HA0A0A0: SetPixelV lHDC, 3, tmph + 15, &H898989: SetPixelV lHDC, 4, tmph + 15, &HDEDEDE: SetPixelV lHDC, 5, tmph + 15, &HE9E9E9: SetPixelV lHDC, 6, tmph + 15, &HEEEEEE: SetPixelV lHDC, 7, tmph + 15, &HF4F4F4: SetPixelV lHDC, 8, tmph + 15, &HF5F5F5: SetPixelV lHDC, 9, tmph + 15, &HFAFAFA: SetPixelV lHDC, 10, tmph + 15, &HFFFDFD: SetPixelV lHDC, 11, tmph + 15, &HFFFEFE: SetPixelV lHDC, 12, tmph + 15, &HFFFDFD: SetPixelV lHDC, 13, tmph + 15, &HFFFEFE: SetPixelV lHDC, 14, tmph + 15, &HFFFDFD: SetPixelV lHDC, 15, tmph + 15, &HFFFEFE: SetPixelV lHDC, 16, tmph + 15, &HFFFDFD: SetPixelV lHDC, 17, tmph + 15, &HFFFEFE
  SetPixelV lHDC, 1, tmph + 16, &HF6F6F6: SetPixelV lHDC, 2, tmph + 16, &HD6D6D6: SetPixelV lHDC, 3, tmph + 16, &H7B7B7B: SetPixelV lHDC, 4, tmph + 16, &H8D8D8D: SetPixelV lHDC, 5, tmph + 16, &HE4E4E4: SetPixelV lHDC, 6, tmph + 16, &HF0F0F0: SetPixelV lHDC, 7, tmph + 16, &HF6F6F6: SetPixelV lHDC, 8, tmph + 16, &HFEFEFE: SetPixelV lHDC, 9, tmph + 16, &HFEFEFE: SetPixelV lHDC, 10, tmph + 16, &HFFFEFE: SetPixelV lHDC, 12, tmph + 16, &HFFFEFE: SetPixelV lHDC, 14, tmph + 16, &HFFFEFE: SetPixelV lHDC, 16, tmph + 16, &HFFFEFE
  SetPixelV lHDC, 1, tmph + 17, &HFDFDFD: SetPixelV lHDC, 2, tmph + 17, &HEDEDED: SetPixelV lHDC, 3, tmph + 17, &HBEBEBE: SetPixelV lHDC, 4, tmph + 17, &H727272: SetPixelV lHDC, 5, tmph + 17, &H898989: SetPixelV lHDC, 6, tmph + 17, &HEBEBEB: SetPixelV lHDC, 7, tmph + 17, &HF5F5F5: SetPixelV lHDC, 8, tmph + 17, &HFCFCFC: SetPixelV lHDC, 10, tmph + 17, &HFDFDFD: SetPixelV lHDC, 11, tmph + 17, &HFDFDFD: SetPixelV lHDC, 12, tmph + 17, &HFDFDFD: SetPixelV lHDC, 13, tmph + 17, &HFDFDFD: SetPixelV lHDC, 14, tmph + 17, &HFDFDFD: SetPixelV lHDC, 15, tmph + 17, &HFDFDFD: SetPixelV lHDC, 16, tmph + 17, &HFDFDFD: SetPixelV lHDC, 17, tmph + 17, &HFDFDFD
  SetPixelV lHDC, 2, tmph + 18, &HF9F9F9: SetPixelV lHDC, 3, tmph + 18, &HE6E6E6: SetPixelV lHDC, 4, tmph + 18, &HB9B9B9: SetPixelV lHDC, 5, tmph + 18, &H717171: SetPixelV lHDC, 6, tmph + 18, &H787878: SetPixelV lHDC, 7, tmph + 18, &HB6B6B6: SetPixelV lHDC, 8, tmph + 18, &HF7F7F7: SetPixelV lHDC, 9, tmph + 18, &HFCFCFC: SetPixelV lHDC, 10, tmph + 18, &HFEFEFE: SetPixelV lHDC, 11, tmph + 18, &HFEFEFE: SetPixelV lHDC, 12, tmph + 18, &HFEFEFE: SetPixelV lHDC, 13, tmph + 18, &HFEFEFE: SetPixelV lHDC, 14, tmph + 18, &HFEFEFE: SetPixelV lHDC, 15, tmph + 18, &HFEFEFE: SetPixelV lHDC, 16, tmph + 18, &HFEFEFE: SetPixelV lHDC, 17, tmph + 18, &HFEFEFE
  SetPixelV lHDC, 2, tmph + 19, &HFEFEFE: SetPixelV lHDC, 3, tmph + 19, &HF8F8F8: SetPixelV lHDC, 4, tmph + 19, &HE6E6E6: SetPixelV lHDC, 5, tmph + 19, &HC8C8C8: SetPixelV lHDC, 6, tmph + 19, &H8E8E8E: SetPixelV lHDC, 7, tmph + 19, &H6C6C6C: SetPixelV lHDC, 8, tmph + 19, &H757575: SetPixelV lHDC, 9, tmph + 19, &H9F9F9F: SetPixelV lHDC, 10, tmph + 19, &HC7C7C7: SetPixelV lHDC, 11, tmph + 19, &HE9E9E9: SetPixelV lHDC, 12, tmph + 19, &HFBFBFB: SetPixelV lHDC, 13, tmph + 19, &HFBFBFB: SetPixelV lHDC, 14, tmph + 19, &HFBFBFB: SetPixelV lHDC, 15, tmph + 19, &HFBFBFB: SetPixelV lHDC, 16, tmph + 19, &HFBFBFB: SetPixelV lHDC, 17, tmph + 19, &HFBFBFB
  SetPixelV lHDC, 3, tmph + 20, &HFEFEFE: SetPixelV lHDC, 4, tmph + 20, &HF9F9F9: SetPixelV lHDC, 5, tmph + 20, &HECECEC: SetPixelV lHDC, 6, tmph + 20, &HDADADA: SetPixelV lHDC, 7, tmph + 20, &HC1C1C1: SetPixelV lHDC, 8, tmph + 20, &H9D9D9D: SetPixelV lHDC, 9, tmph + 20, &H7B7B7B: SetPixelV lHDC, 10, tmph + 20, &H5E5E5E: SetPixelV lHDC, 11, tmph + 20, &H535353: SetPixelV lHDC, 12, tmph + 20, &H4D4D4D: SetPixelV lHDC, 13, tmph + 20, &H4B4B4B: SetPixelV lHDC, 14, tmph + 20, &H505050: SetPixelV lHDC, 15, tmph + 20, &H525252: SetPixelV lHDC, 16, tmph + 20, &H555555: SetPixelV lHDC, 17, tmph + 20, &H545454
  SetPixelV lHDC, 5, tmph + 21, &HFCFCFC: SetPixelV lHDC, 6, tmph + 21, &HF5F5F5: SetPixelV lHDC, 7, tmph + 21, &HEBEBEB: SetPixelV lHDC, 8, tmph + 21, &HE1E1E1: SetPixelV lHDC, 9, tmph + 21, &HD6D6D6: SetPixelV lHDC, 10, tmph + 21, &HCECECE: SetPixelV lHDC, 11, tmph + 21, &HC9C9C9: SetPixelV lHDC, 12, tmph + 21, &HC7C7C7: SetPixelV lHDC, 13, tmph + 21, &HC7C7C7: SetPixelV lHDC, 14, tmph + 21, &HC6C6C6: SetPixelV lHDC, 15, tmph + 21, &HC6C6C6: SetPixelV lHDC, 16, tmph + 21, &HC5C5C5: SetPixelV lHDC, 17, tmph + 21, &HC5C5C5
  SetPixelV lHDC, 7, tmph + 22, &HFDFDFD: SetPixelV lHDC, 8, tmph + 22, &HF9F9F9: SetPixelV lHDC, 9, tmph + 22, &HF4F4F4: SetPixelV lHDC, 10, tmph + 22, &HF0F0F0: SetPixelV lHDC, 11, tmph + 22, &HEEEEEE: SetPixelV lHDC, 12, tmph + 22, &HEDEDED: SetPixelV lHDC, 13, tmph + 22, &HECECEC: SetPixelV lHDC, 14, tmph + 22, &HECECEC: SetPixelV lHDC, 15, tmph + 22, &HECECEC: SetPixelV lHDC, 16, tmph + 22, &HECECEC: SetPixelV lHDC, 17, tmph + 22, &HECECEC
  tmpw = lw - 34
  SetPixelV lHDC, tmpw + 17, 0, &H67696A: SetPixelV lHDC, tmpw + 18, 0, &H666869: SetPixelV lHDC, tmpw + 19, 0, &H716F6F: SetPixelV lHDC, tmpw + 20, 0, &H6F6D6D: SetPixelV lHDC, tmpw + 21, 0, &H6F706E: SetPixelV lHDC, tmpw + 22, 0, &H727371: SetPixelV lHDC, tmpw + 23, 0, &H6E6E6E: SetPixelV lHDC, tmpw + 24, 0, &H707070: SetPixelV lHDC, tmpw + 25, 0, &HA6A6A6: SetPixelV lHDC, tmpw + 26, 0, &HEEEEEE: SetPixelV lHDC, tmpw + 34, 0, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 1, &HF5F4F6: SetPixelV lHDC, tmpw + 18, 1, &HF5F4F6: SetPixelV lHDC, tmpw + 19, 1, &HF5F4F6: SetPixelV lHDC, tmpw + 20, 1, &HF5F4F6: SetPixelV lHDC, tmpw + 21, 1, &HF4F3F5: SetPixelV lHDC, tmpw + 22, 1, &HF1F0F2: SetPixelV lHDC, tmpw + 23, 1, &HE0E0E0: SetPixelV lHDC, tmpw + 24, 1, &HC3C3C3: SetPixelV lHDC, tmpw + 25, 1, &H848484: SetPixelV lHDC, tmpw + 26, 1, &H6B6B6B: SetPixelV lHDC, tmpw + 27, 1, &HA0A0A0: SetPixelV lHDC, tmpw + 28, 1, &HF7F7F7: SetPixelV lHDC, tmpw + 34, 1, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 2, &HF3F2F4: SetPixelV lHDC, tmpw + 18, 2, &HF2F1F3: SetPixelV lHDC, tmpw + 19, 2, &HF3F2F4: SetPixelV lHDC, tmpw + 20, 2, &HF3F2F4: SetPixelV lHDC, tmpw + 21, 2, &HF0EFF1: SetPixelV lHDC, tmpw + 22, 2, &HF2F1F3: SetPixelV lHDC, tmpw + 23, 2, &HF6F6F6: SetPixelV lHDC, tmpw + 24, 2, &HE8E8E8: SetPixelV lHDC, tmpw + 25, 2, &HE0E0E0: SetPixelV lHDC, tmpw + 26, 2, &H999999: SetPixelV lHDC, tmpw + 27, 2, &H696969: SetPixelV lHDC, tmpw + 28, 2, &H717171: SetPixelV lHDC, tmpw + 29, 2, &HEBEBEB: SetPixelV lHDC, tmpw + 34, 2, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 3, &HEEEEEE: SetPixelV lHDC, tmpw + 18, 3, &HEDEDED: SetPixelV lHDC, tmpw + 19, 3, &HEEEEEE: SetPixelV lHDC, tmpw + 20, 3, &HEEEEEE: SetPixelV lHDC, tmpw + 21, 3, &HEEEEEE: SetPixelV lHDC, tmpw + 22, 3, &HEEEEEE: SetPixelV lHDC, tmpw + 23, 3, &HE9E9E9: SetPixelV lHDC, tmpw + 24, 3, &HEAEAEA: SetPixelV lHDC, tmpw + 25, 3, &HE7E7E7: SetPixelV lHDC, tmpw + 26, 3, &HD0D0D0: SetPixelV lHDC, tmpw + 27, 3, &H939393: SetPixelV lHDC, tmpw + 28, 3, &H727272: SetPixelV lHDC, tmpw + 29, 3, &H6F6F6F: SetPixelV lHDC, tmpw + 30, 3, &HEFEFEF: SetPixelV lHDC, tmpw + 34, 3, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 4, &HEBEBEB: SetPixelV lHDC, tmpw + 18, 4, &HEBEBEB: SetPixelV lHDC, tmpw + 19, 4, &HEBEBEB: SetPixelV lHDC, tmpw + 20, 4, &HEBEBEB: SetPixelV lHDC, tmpw + 21, 4, &HEDEDED: SetPixelV lHDC, tmpw + 22, 4, &HE6E6E6: SetPixelV lHDC, tmpw + 23, 4, &HE9E9E9: SetPixelV lHDC, tmpw + 24, 4, &HE6E6E6: SetPixelV lHDC, tmpw + 25, 4, &HDEDEDE: SetPixelV lHDC, tmpw + 26, 4, &HDCDCDC: SetPixelV lHDC, tmpw + 27, 4, &HB2B2B2: SetPixelV lHDC, tmpw + 28, 4, &H919191: SetPixelV lHDC, tmpw + 29, 4, &H6E6E6E: SetPixelV lHDC, tmpw + 30, 4, &H7F7F7F: SetPixelV lHDC, tmpw + 31, 4, &HFAFAFA: SetPixelV lHDC, tmpw + 34, 4, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 5, &HEBE8EA: SetPixelV lHDC, tmpw + 18, 5, &HEAE7E9: SetPixelV lHDC, tmpw + 19, 5, &HEBE8EA: SetPixelV lHDC, tmpw + 20, 5, &HEBE8EA: SetPixelV lHDC, tmpw + 21, 5, &HE5E8E6: SetPixelV lHDC, tmpw + 22, 5, &HE7EAE8: SetPixelV lHDC, tmpw + 23, 5, &HE5E5E5: SetPixelV lHDC, tmpw + 24, 5, &HE3E3E3: SetPixelV lHDC, tmpw + 25, 5, &HDFDFDF: SetPixelV lHDC, tmpw + 26, 5, &HDCDCDC: SetPixelV lHDC, tmpw + 27, 5, &HC3C3C3: SetPixelV lHDC, tmpw + 28, 5, &HA7A7A7: SetPixelV lHDC, tmpw + 29, 5, &H969696: SetPixelV lHDC, tmpw + 30, 5, &H717171: SetPixelV lHDC, tmpw + 31, 5, &HC5C5C5: SetPixelV lHDC, tmpw + 32, 5, &HFEFEFE: SetPixelV lHDC, tmpw + 34, 5, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 6, &HEBE8EA: SetPixelV lHDC, tmpw + 18, 6, &HEBE8EA: SetPixelV lHDC, tmpw + 19, 6, &HEBE8EA: SetPixelV lHDC, tmpw + 20, 6, &HEBE8EA: SetPixelV lHDC, tmpw + 21, 6, &HE8EBE9: SetPixelV lHDC, tmpw + 22, 6, &HE3E6E4: SetPixelV lHDC, tmpw + 23, 6, &HE5E5E5: SetPixelV lHDC, tmpw + 24, 6, &HE2E2E2: SetPixelV lHDC, tmpw + 25, 6, &HE0E0E0: SetPixelV lHDC, tmpw + 26, 6, &HDADADA: SetPixelV lHDC, tmpw + 27, 6, &HC7C7C7: SetPixelV lHDC, tmpw + 28, 6, &HB5B5B5: SetPixelV lHDC, tmpw + 29, 6, &HA6A6A6: SetPixelV lHDC, tmpw + 30, 6, &H8C8C8C: SetPixelV lHDC, tmpw + 31, 6, &H808080: SetPixelV lHDC, tmpw + 32, 6, &HF8F8F8: SetPixelV lHDC, tmpw + 34, 6, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 7, &HEAEAEA: SetPixelV lHDC, tmpw + 18, 7, &HEAEAEA: SetPixelV lHDC, tmpw + 19, 7, &HEAEAEA: SetPixelV lHDC, tmpw + 20, 7, &HEAEAEA: SetPixelV lHDC, tmpw + 21, 7, &HE9E6E8: SetPixelV lHDC, tmpw + 22, 7, &HE9E6E8: SetPixelV lHDC, tmpw + 23, 7, &HE4E4E4: SetPixelV lHDC, tmpw + 24, 7, &HE2E2E2: SetPixelV lHDC, tmpw + 25, 7, &HDFDFDF: SetPixelV lHDC, tmpw + 26, 7, &HD7D7D7: SetPixelV lHDC, tmpw + 27, 7, &HC4C4C4: SetPixelV lHDC, tmpw + 28, 7, &HB7B7B7: SetPixelV lHDC, tmpw + 29, 7, &HB4B5B3: SetPixelV lHDC, tmpw + 30, 7, &H9D9E9C: SetPixelV lHDC, tmpw + 31, 7, &H777777: SetPixelV lHDC, tmpw + 32, 7, &HE7E7E7: SetPixelV lHDC, tmpw + 34, 7, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 8, &HE1E1E1: SetPixelV lHDC, tmpw + 18, 8, &HE0E0E0: SetPixelV lHDC, tmpw + 19, 8, &HE1E1E1: SetPixelV lHDC, tmpw + 20, 8, &HE1E1E1: SetPixelV lHDC, tmpw + 21, 8, &HDFDCDE: SetPixelV lHDC, tmpw + 22, 8, &HDDDADC: SetPixelV lHDC, tmpw + 23, 8, &HDBDBDB: SetPixelV lHDC, tmpw + 24, 8, &HD6D6D6: SetPixelV lHDC, tmpw + 25, 8, &HD5D5D5: SetPixelV lHDC, tmpw + 26, 8, &HD1D1D1: SetPixelV lHDC, tmpw + 27, 8, &HC9C9C9: SetPixelV lHDC, tmpw + 28, 8, &HC4C4C4: SetPixelV lHDC, tmpw + 29, 8, &HC0C1BF: SetPixelV lHDC, tmpw + 30, 8, &HAFB0AE: SetPixelV lHDC, tmpw + 31, 8, &H818181: SetPixelV lHDC, tmpw + 32, 8, &HC3C3C3: SetPixelV lHDC, tmpw + 33, 8, &HFDFDFD: SetPixelV lHDC, tmpw + 34, 8, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 9, &HE5E2E3: SetPixelV lHDC, tmpw + 18, 9, &HE5E2E3: SetPixelV lHDC, tmpw + 19, 9, &HE5E2E3: SetPixelV lHDC, tmpw + 20, 9, &HE5E2E3: SetPixelV lHDC, tmpw + 21, 9, &HE1E1E1: SetPixelV lHDC, tmpw + 22, 9, &HE1E1E1: SetPixelV lHDC, tmpw + 23, 9, &HE1E1E1: SetPixelV lHDC, tmpw + 24, 9, &HDDDDDD: SetPixelV lHDC, tmpw + 25, 9, &HDBDBDB: SetPixelV lHDC, tmpw + 26, 9, &HD8D8D8: SetPixelV lHDC, tmpw + 27, 9, &HD2D2D2: SetPixelV lHDC, tmpw + 28, 9, &HCBCBCB: SetPixelV lHDC, tmpw + 29, 9, &HC4C4C4: SetPixelV lHDC, tmpw + 30, 9, &HBABABA: SetPixelV lHDC, tmpw + 31, 9, &H989898: SetPixelV lHDC, tmpw + 32, 9, &HA6A6A6: SetPixelV lHDC, tmpw + 33, 9, &HF9F9F9: SetPixelV lHDC, tmpw + 34, 9, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 10, &HEAE7E8: SetPixelV lHDC, tmpw + 18, 10, &HEAE7E8: SetPixelV lHDC, tmpw + 19, 10, &HEAE7E8: SetPixelV lHDC, tmpw + 20, 10, &HEAE7E8: SetPixelV lHDC, tmpw + 21, 10, &HE7E7E7: SetPixelV lHDC, tmpw + 22, 10, &HE6E6E6: SetPixelV lHDC, tmpw + 23, 10, &HE4E4E4: SetPixelV lHDC, tmpw + 24, 10, &HE0E0E0: SetPixelV lHDC, tmpw + 25, 10, &HE0E0E0: SetPixelV lHDC, tmpw + 26, 10, &HDEDEDE: SetPixelV lHDC, tmpw + 27, 10, &HD9D9D9: SetPixelV lHDC, tmpw + 28, 10, &HD3D3D3: SetPixelV lHDC, tmpw + 29, 10, &HCCCCCC: SetPixelV lHDC, tmpw + 30, 10, &HC3C3C3: SetPixelV lHDC, tmpw + 31, 10, &HA3A3A3: SetPixelV lHDC, tmpw + 32, 10, &H9C9C9C: SetPixelV lHDC, tmpw + 33, 10, &HF6F6F6: SetPixelV lHDC, tmpw + 34, 10, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 11, &HE8EBE9: SetPixelV lHDC, tmpw + 18, 11, &HE8EBE9: SetPixelV lHDC, tmpw + 19, 11, &HE8EBE9: SetPixelV lHDC, tmpw + 20, 11, &HE8EBE9: SetPixelV lHDC, tmpw + 21, 11, &HE9EAE8: SetPixelV lHDC, tmpw + 22, 11, &HE8E9E7: SetPixelV lHDC, tmpw + 23, 11, &HE9E9E9: SetPixelV lHDC, tmpw + 24, 11, &HE5E5E5: SetPixelV lHDC, tmpw + 25, 11, &HE4E4E4: SetPixelV lHDC, tmpw + 26, 11, &HE2E2E2: SetPixelV lHDC, tmpw + 27, 11, &HDBDBDB: SetPixelV lHDC, tmpw + 28, 11, &HD9D9D9: SetPixelV lHDC, tmpw + 29, 11, &HD1D1D1: SetPixelV lHDC, tmpw + 30, 11, &HC8C8C8: SetPixelV lHDC, tmpw + 31, 11, &HA4A4A4: SetPixelV lHDC, tmpw + 32, 11, &HA2A2A2: SetPixelV lHDC, tmpw + 33, 11, &HF4F4F4: SetPixelV lHDC, tmpw + 34, 11, &HFFFFFFFF
  SetPixelV lHDC, tmpw + 17, 12, &HEEF1EF: SetPixelV lHDC, tmpw + 18, 12, &HEEF1EF: SetPixelV lHDC, tmpw + 19, 12, &HEEF1EF: SetPixelV lHDC, tmpw + 20, 12, &HEEF1EF: SetPixelV lHDC, tmpw + 21, 12, &HEEEFED: SetPixelV lHDC, tmpw + 22, 12, &HEFF0EE: SetPixelV lHDC, tmpw + 23, 12, &HEEEEEE: SetPixelV lHDC, tmpw + 24, 12, &HECECEC: SetPixelV lHDC, tmpw + 25, 12, &HEAEAEA: SetPixelV lHDC, tmpw + 26, 12, &HE7E7E7: SetPixelV lHDC, tmpw + 27, 12, &HE2E2E2: SetPixelV lHDC, tmpw + 28, 12, &HDFDFDF: SetPixelV lHDC, tmpw + 29, 12, &HD8D8D8: SetPixelV lHDC, tmpw + 30, 12, &HD4D4D4: SetPixelV lHDC, tmpw + 31, 12, &H999999: SetPixelV lHDC, tmpw + 32, 12, &HAFAFAF: SetPixelV lHDC, tmpw + 33, 12, &HF5F5F5: SetPixelV lHDC, tmpw + 34, 12, &HFFFFFFFF
  tmph = lh - 22
  tmpw = lw - 34
  SetPixelV lHDC, tmpw + 17, tmph + 12, &HEEF1EF: SetPixelV lHDC, tmpw + 18, tmph + 12, &HEEF1EF: SetPixelV lHDC, tmpw + 19, tmph + 12, &HEEF1EF: SetPixelV lHDC, tmpw + 20, tmph + 12, &HEEF1EF: SetPixelV lHDC, tmpw + 21, tmph + 12, &HEEEFED: SetPixelV lHDC, tmpw + 22, tmph + 12, &HEFF0EE: SetPixelV lHDC, tmpw + 23, tmph + 12, &HEEEEEE: SetPixelV lHDC, tmpw + 24, tmph + 12, &HECECEC: SetPixelV lHDC, tmpw + 25, tmph + 12, &HEAEAEA: SetPixelV lHDC, tmpw + 26, tmph + 12, &HE7E7E7: SetPixelV lHDC, tmpw + 27, tmph + 12, &HE2E2E2: SetPixelV lHDC, tmpw + 28, tmph + 12, &HDFDFDF: SetPixelV lHDC, tmpw + 29, tmph + 12, &HD8D8D8: SetPixelV lHDC, tmpw + 30, tmph + 12, &HD4D4D4: SetPixelV lHDC, tmpw + 31, tmph + 12, &H999999: SetPixelV lHDC, tmpw + 32, tmph + 12, &HAFAFAF: SetPixelV lHDC, tmpw + 33, tmph + 12, &HF5F5F5
  SetPixelV lHDC, tmpw + 17, tmph + 13, &HF2F2F2: SetPixelV lHDC, tmpw + 18, tmph + 13, &HF2F2F2: SetPixelV lHDC, tmpw + 19, tmph + 13, &HF2F2F2: SetPixelV lHDC, tmpw + 20, tmph + 13, &HF2F2F2: SetPixelV lHDC, tmpw + 21, tmph + 13, &HF5F4F6: SetPixelV lHDC, tmpw + 22, tmph + 13, &HF0EFF1: SetPixelV lHDC, tmpw + 23, tmph + 13, &HF2F2F2: SetPixelV lHDC, tmpw + 24, tmph + 13, &HF2F2F2: SetPixelV lHDC, tmpw + 25, tmph + 13, &HECECEC: SetPixelV lHDC, tmpw + 26, tmph + 13, &HEAEAEA: SetPixelV lHDC, tmpw + 27, tmph + 13, &HEBEBEB: SetPixelV lHDC, tmpw + 28, tmph + 13, &HE3E3E3: SetPixelV lHDC, tmpw + 29, tmph + 13, &HDEDEDE: SetPixelV lHDC, tmpw + 30, tmph + 13, &HD1D1D1: SetPixelV lHDC, tmpw + 31, tmph + 13, &H8A8A8A: SetPixelV lHDC, tmpw + 32, tmph + 13, &HD5D5D5: SetPixelV lHDC, tmpw + 33, tmph + 13, &HF8F8F8
  SetPixelV lHDC, tmpw + 17, tmph + 14, &HF5F5F5: SetPixelV lHDC, tmpw + 18, tmph + 14, &HF5F5F5: SetPixelV lHDC, tmpw + 19, tmph + 14, &HF5F5F5: SetPixelV lHDC, tmpw + 20, tmph + 14, &HF5F5F5: SetPixelV lHDC, tmpw + 21, tmph + 14, &HF8F7F9: SetPixelV lHDC, tmpw + 22, tmph + 14, &HF7F6F8: SetPixelV lHDC, tmpw + 23, tmph + 14, &HF7F7F7: SetPixelV lHDC, tmpw + 24, tmph + 14, &HF5F5F5: SetPixelV lHDC, tmpw + 25, tmph + 14, &HEFEFEF: SetPixelV lHDC, tmpw + 26, tmph + 14, &HEEEEEE: SetPixelV lHDC, tmpw + 27, tmph + 14, &HECECEC: SetPixelV lHDC, tmpw + 28, tmph + 14, &HE5E5E5: SetPixelV lHDC, tmpw + 29, tmph + 14, &HDEDEDE: SetPixelV lHDC, tmpw + 30, tmph + 14, &HB3B3B3: SetPixelV lHDC, tmpw + 31, tmph + 14, &H808080: SetPixelV lHDC, tmpw + 32, tmph + 14, &HE8E8E8: SetPixelV lHDC, tmpw + 33, tmph + 14, &HFDFDFD
  SetPixelV lHDC, tmpw + 17, tmph + 15, &HFFFEFE: SetPixelV lHDC, tmpw + 18, tmph + 15, &HFFFDFD: SetPixelV lHDC, tmpw + 19, tmph + 15, &HFFFEFE: SetPixelV lHDC, tmpw + 20, tmph + 15, &HFFFEFE: SetPixelV lHDC, tmpw + 21, tmph + 15, &HFBFBFB: SetPixelV lHDC, tmpw + 22, tmph + 15, &HFCFCFC: SetPixelV lHDC, tmpw + 23, tmph + 15, &HFEFEFE: SetPixelV lHDC, tmpw + 24, tmph + 15, &HF8F8F8: SetPixelV lHDC, tmpw + 25, tmph + 15, &HF7F7F7: SetPixelV lHDC, tmpw + 26, tmph + 15, &HF5F5F5: SetPixelV lHDC, tmpw + 27, tmph + 15, &HEDEDED: SetPixelV lHDC, tmpw + 28, tmph + 15, &HEAEAEA: SetPixelV lHDC, tmpw + 29, tmph + 15, &HE0E0E0: SetPixelV lHDC, tmpw + 30, tmph + 15, &H8D8D8D: SetPixelV lHDC, tmpw + 31, tmph + 15, &HBABABA: SetPixelV lHDC, tmpw + 32, tmph + 15, &HF1F1F1
  SetPixelV lHDC, tmpw + 18, tmph + 16, &HFFFEFE: SetPixelV lHDC, tmpw + 22, tmph + 16, &HFEFEFE: SetPixelV lHDC, tmpw + 23, tmph + 16, &HFEFEFE: SetPixelV lHDC, tmpw + 25, tmph + 16, &HFCFCFC: SetPixelV lHDC, tmpw + 26, tmph + 16, &HF6F6F6: SetPixelV lHDC, tmpw + 27, tmph + 16, &HF2F2F2: SetPixelV lHDC, tmpw + 28, tmph + 16, &HE7E7E7: SetPixelV lHDC, tmpw + 29, tmph + 16, &H989898: SetPixelV lHDC, tmpw + 30, tmph + 16, &H828282: SetPixelV lHDC, tmpw + 31, tmph + 16, &HE2E2E2: SetPixelV lHDC, tmpw + 32, tmph + 16, &HF9F9F9
  SetPixelV lHDC, tmpw + 17, tmph + 17, &HFDFDFD: SetPixelV lHDC, tmpw + 18, tmph + 17, &HFDFDFD: SetPixelV lHDC, tmpw + 19, tmph + 17, &HFDFDFD: SetPixelV lHDC, tmpw + 20, tmph + 17, &HFDFDFD: SetPixelV lHDC, tmpw + 21, tmph + 17, &HFEFEFE: SetPixelV lHDC, tmpw + 23, tmph + 17, &HFEFEFE: SetPixelV lHDC, tmpw + 25, tmph + 17, &HFEFEFE: SetPixelV lHDC, tmpw + 26, tmph + 17, &HF6F6F6: SetPixelV lHDC, tmpw + 27, tmph + 17, &HF1F1F1: SetPixelV lHDC, tmpw + 28, tmph + 17, &H979797: SetPixelV lHDC, tmpw + 29, tmph + 17, &H6F6F6F: SetPixelV lHDC, tmpw + 30, tmph + 17, &HD2D2D2: SetPixelV lHDC, tmpw + 31, tmph + 17, &HF2F2F2: SetPixelV lHDC, tmpw + 32, tmph + 17, &HFEFEFE
  SetPixelV lHDC, tmpw + 17, tmph + 18, &HFEFEFE: SetPixelV lHDC, tmpw + 18, tmph + 18, &HFEFEFE: SetPixelV lHDC, tmpw + 19, tmph + 18, &HFEFEFE: SetPixelV lHDC, tmpw + 20, tmph + 18, &HFEFEFE: SetPixelV lHDC, tmpw + 22, tmph + 18, &HFDFDFD: SetPixelV lHDC, tmpw + 23, tmph + 18, &HFEFEFE: SetPixelV lHDC, tmpw + 24, tmph + 18, &HFDFDFD: SetPixelV lHDC, tmpw + 25, tmph + 18, &HFCFCFC: SetPixelV lHDC, tmpw + 26, tmph + 18, &HC5C5C5: SetPixelV lHDC, tmpw + 27, tmph + 18, &H838383: SetPixelV lHDC, tmpw + 28, tmph + 18, &H6F6F6F: SetPixelV lHDC, tmpw + 29, tmph + 18, &HC8C8C8: SetPixelV lHDC, tmpw + 30, tmph + 18, &HEBEBEB: SetPixelV lHDC, tmpw + 31, tmph + 18, &HFCFCFC
  SetPixelV lHDC, tmpw + 17, tmph + 19, &HFBFBFB: SetPixelV lHDC, tmpw + 18, tmph + 19, &HFBFBFB: SetPixelV lHDC, tmpw + 19, tmph + 19, &HFBFBFB: SetPixelV lHDC, tmpw + 20, tmph + 19, &HFBFBFB: SetPixelV lHDC, tmpw + 21, tmph + 19, &HFAFAFA: SetPixelV lHDC, tmpw + 22, tmph + 19, &HEFEFEF: SetPixelV lHDC, tmpw + 23, tmph + 19, &HD0D0D0: SetPixelV lHDC, tmpw + 24, tmph + 19, &HA3A3A3: SetPixelV lHDC, tmpw + 25, tmph + 19, &H7E7E7E: SetPixelV lHDC, tmpw + 26, tmph + 19, &H6A6A6A: SetPixelV lHDC, tmpw + 27, tmph + 19, &H8F8F8F: SetPixelV lHDC, tmpw + 28, tmph + 19, &HCDCDCD: SetPixelV lHDC, tmpw + 29, tmph + 19, &HE8E8E8: SetPixelV lHDC, tmpw + 30, tmph + 19, &HFAFAFA
  SetPixelV lHDC, tmpw + 17, tmph + 20, &H545454: SetPixelV lHDC, tmpw + 18, tmph + 20, &H555555: SetPixelV lHDC, tmpw + 19, tmph + 20, &H525252: SetPixelV lHDC, tmpw + 20, tmph + 20, &H505050: SetPixelV lHDC, tmpw + 21, tmph + 20, &H535353: SetPixelV lHDC, tmpw + 22, tmph + 20, &H525252: SetPixelV lHDC, tmpw + 23, tmph + 20, &H616161: SetPixelV lHDC, tmpw + 24, tmph + 20, &H7A7A7A: SetPixelV lHDC, tmpw + 25, tmph + 20, &HA3A3A3: SetPixelV lHDC, tmpw + 26, tmph + 20, &HC5C5C5: SetPixelV lHDC, tmpw + 27, tmph + 20, &HDADADA: SetPixelV lHDC, tmpw + 28, tmph + 20, &HEDEDED: SetPixelV lHDC, tmpw + 29, tmph + 20, &HFAFAFA
  SetPixelV lHDC, tmpw + 17, tmph + 21, &HC5C5C5: SetPixelV lHDC, tmpw + 18, tmph + 21, &HC5C5C5: SetPixelV lHDC, tmpw + 19, tmph + 21, &HC6C6C6: SetPixelV lHDC, tmpw + 20, tmph + 21, &HC6C6C6: SetPixelV lHDC, tmpw + 21, tmph + 21, &HC6C6C6: SetPixelV lHDC, tmpw + 22, tmph + 21, &HC9C9C9: SetPixelV lHDC, tmpw + 23, tmph + 21, &HCECECE: SetPixelV lHDC, tmpw + 24, tmph + 21, &HD7D7D7: SetPixelV lHDC, tmpw + 25, tmph + 21, &HE1E1E1: SetPixelV lHDC, tmpw + 26, tmph + 21, &HECECEC: SetPixelV lHDC, tmpw + 27, tmph + 21, &HF6F6F6: SetPixelV lHDC, tmpw + 28, tmph + 21, &HFDFDFD
  SetPixelV lHDC, tmpw + 17, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 18, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 19, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 20, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 21, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 22, tmph + 22, &HEDEDED: SetPixelV lHDC, tmpw + 23, tmph + 22, &HF0F0F0: SetPixelV lHDC, tmpw + 24, tmph + 22, &HF4F4F4: SetPixelV lHDC, tmpw + 25, tmph + 22, &HFAFAFA: SetPixelV lHDC, tmpw + 26, tmph + 22, &HFDFDFD
  'Vlines
  tmph = 11:     tmph1 = lh - 10:     tmpw = lw - 34
  APILine 0, tmph, 0, tmph1, &HF7F7F7: APILine 1, tmph, 1, tmph1, &HA0A0A0: APILine 2, tmph, 2, tmph1, &H999999: APILine 3, tmph, 3, tmph1, &HC3C3C3
  APILine 4, tmph, 4, tmph1, &HC9C9C9: APILine 5, tmph, 5, tmph1, &HD5D5D5: APILine 6, tmph, 6, tmph1, &HD7D7D7: APILine 7, tmph, 7, tmph1, &HDFDFDF
  APILine 8, tmph, 8, tmph1, &HE0E0E0: APILine 9, tmph, 9, tmph1, &HE0E0E0: APILine 10, tmph, 10, tmph1, &HE4E4E4: APILine 11, tmph, 11, tmph1, &HE6E8E6
  APILine 12, tmph, 12, tmph1, &HE8E7E7: APILine 13, tmph, 13, tmph1, &HEAE7E8: APILine 14, tmph, 14, tmph1, &HEAE7E8: APILine 15, tmph, 15, tmph1, &HEAE7E8
  APILine 16, tmph, 16, tmph1, &HEAE7E8: APILine 17, tmph, 17, tmph1, &HEAE7E8: APILine tmpw + 17, tmph, tmpw + 17, tmph1, &HEAE7E8: APILine tmpw + 18, tmph, tmpw + 18, tmph1, &HEAE7E8
  APILine tmpw + 19, tmph, tmpw + 19, tmph1, &HEAE7E8: APILine tmpw + 20, tmph, tmpw + 20, tmph1, &HEAE7E8: APILine tmpw + 21, tmph, tmpw + 21, tmph1, &HE7E7E7
  APILine tmpw + 22, tmph, tmpw + 22, tmph1, &HE6E6E6: APILine tmpw + 23, tmph, tmpw + 23, tmph1, &HE4E4E4: APILine tmpw + 24, tmph, tmpw + 24, tmph1, &HE0E0E0
  APILine tmpw + 25, tmph, tmpw + 25, tmph1, &HE0E0E0: APILine tmpw + 26, tmph, tmpw + 26, tmph1, &HDEDEDE: APILine tmpw + 27, tmph, tmpw + 27, tmph1, &HD9D9D9
  APILine tmpw + 28, tmph, tmpw + 28, tmph1, &HD3D3D3: APILine tmpw + 29, tmph, tmpw + 29, tmph1, &HCCCCCC: APILine tmpw + 30, tmph, tmpw + 30, tmph1, &HC3C3C3
  APILine tmpw + 31, tmph, tmpw + 31, tmph1, &HA3A3A3: APILine tmpw + 32, tmph, tmpw + 32, tmph1, &H9C9C9C: APILine tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6
  'HLines
  APILine 17, 0, lw - 17, 0, &H67696A
  APILine 17, 1, lw - 17, 1, &HF5F4F6
  APILine 17, 2, lw - 17, 2, &HF3F2F4
  APILine 17, 3, lw - 17, 3, &HEEEEEE
  APILine 17, 4, lw - 17, 4, &HEBEBEB
  APILine 17, 5, lw - 17, 5, &HEBE8EA
  APILine 17, 6, lw - 17, 6, &HEBE8EA
  APILine 17, 7, lw - 17, 7, &HEAEAEA
  APILine 17, 8, lw - 17, 8, &HE1E1E1
  APILine 17, 9, lw - 17, 9, &HE5E2E3
  APILine 17, 10, lw - 17, 10, &HEAE7E8
  APILine 17, 11, lw - 17, 11, &HE8EBE9
  tmph = lh - 22
  APILine 17, tmph + 11, lw - 17, tmph + 11, &HE8EBE9
  APILine 17, tmph + 12, lw - 17, tmph + 12, &HEEF1EF
  APILine 17, tmph + 13, lw - 17, tmph + 13, &HF2F2F2
  APILine 17, tmph + 14, lw - 17, tmph + 14, &HF5F5F5
  APILine 17, tmph + 15, lw - 17, tmph + 15, &HFFFEFE
  APILine 17, tmph + 16, lw - 17, tmph + 16, &HFFFFFF
  APILine 17, tmph + 17, lw - 17, tmph + 17, &HFDFDFD
  APILine 17, tmph + 18, lw - 17, tmph + 18, &HFEFEFE
  APILine 17, tmph + 19, lw - 17, tmph + 19, &HFBFBFB
  APILine 17, tmph + 20, lw - 17, tmph + 20, &H545454
  APILine 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
  APILine 17, tmph + 22, lw - 17, tmph + 22, &HECECEC
  Exit Sub

DrawMacOSXButtonNormal_Error:
End Sub

Private Sub DrawMacOSXButtonHot()
  On Error GoTo DrawMacOSXButtonHot_Error

  Dim lHDC As Long
  lHDC = UserControl.hdc
  'Variable vars (real into code)
  Dim lh As Long, lw As Long
  lh = UserControl.ScaleHeight: lw = UserControl.ScaleWidth
  Dim tmph As Long, tmpw As Long
  Dim tmph1 As Long, tmpw1 As Long
  APIFillRectByCoords hdc, 18, 11, lw - 34, lh - 19, &HE2A66A
  SetPixelV lHDC, 6, 0, &HFEFEFE: SetPixelV lHDC, 7, 0, &HE6E5E5: SetPixelV lHDC, 8, 0, &HA9A5A5: SetPixelV lHDC, 9, 0, &H6C5E5E: SetPixelV lHDC, 10, 0, &H482729: SetPixelV lHDC, 11, 0, &H370D0C: SetPixelV lHDC, 12, 0, &H370706: SetPixelV lHDC, 13, 0, &H360605: SetPixelV lHDC, 14, 0, &H3A0606: SetPixelV lHDC, 15, 0, &H410807: SetPixelV lHDC, 16, 0, &H450707: SetPixelV lHDC, 17, 0, &H450608:
  SetPixelV lHDC, 5, 1, &HF0EFEF: SetPixelV lHDC, 6, 1, &HA38A8C: SetPixelV lHDC, 7, 1, &H6E342F: SetPixelV lHDC, 8, 1, &H661F1A: SetPixelV lHDC, 9, 1, &H9B6A63: SetPixelV lHDC, 10, 1, &HC9A29D: SetPixelV lHDC, 11, 1, &HE2BFBD: SetPixelV lHDC, 12, 1, &HE8C9C6: SetPixelV lHDC, 13, 1, &HEFD3CC: SetPixelV lHDC, 14, 1, &HEFD3CC: SetPixelV lHDC, 15, 1, &HF0D5C9: SetPixelV lHDC, 16, 1, &HF0D5C9: SetPixelV lHDC, 17, 1, &HF1D4C9:
  SetPixelV lHDC, 3, 2, &HFEFEFE: SetPixelV lHDC, 4, 2, &HE5E5E5: SetPixelV lHDC, 5, 2, &H755E5E: SetPixelV lHDC, 6, 2, &H41070C: SetPixelV lHDC, 7, 2, &H7F2D28: SetPixelV lHDC, 8, 2, &HEC9892: SetPixelV lHDC, 9, 2, &HECB6AF: SetPixelV lHDC, 10, 2, &HE3BBB6: SetPixelV lHDC, 11, 2, &HE3C0BD: SetPixelV lHDC, 12, 2, &HE1C2BF: SetPixelV lHDC, 13, 2, &HDFC3BC: SetPixelV lHDC, 14, 2, &HDFC3BC: SetPixelV lHDC, 15, 2, &HE4C9BD: SetPixelV lHDC, 16, 2, &HE4C9BD: SetPixelV lHDC, 17, 2, &HE5C8BD:
  SetPixelV lHDC, 3, 3, &HEEEEEE: SetPixelV lHDC, 4, 3, &H8A5A5A: SetPixelV lHDC, 5, 3, &H7A0702: SetPixelV lHDC, 6, 3, &H901501: SetPixelV lHDC, 7, 3, &HC38365: SetPixelV lHDC, 8, 3, &HE3B08F: SetPixelV lHDC, 9, 3, &HE1B394: SetPixelV lHDC, 10, 3, &HE5B798: SetPixelV lHDC, 11, 3, &HE6BC99: SetPixelV lHDC, 12, 3, &HE7BD9A: SetPixelV lHDC, 13, 3, &HE4BC99: SetPixelV lHDC, 14, 3, &HE7BF9C: SetPixelV lHDC, 15, 3, &HE9C1A1: SetPixelV lHDC, 16, 3, &HE8C0A1: SetPixelV lHDC, 17, 3, &HE8C0A1:
  SetPixelV lHDC, 2, 4, &HFBFBFB: SetPixelV lHDC, 3, 4, &H897879: SetPixelV lHDC, 4, 4, &H4D0909: SetPixelV lHDC, 5, 4, &H951905: SetPixelV lHDC, 6, 4, &HBF422E: SetPixelV lHDC, 7, 4, &HD49475: SetPixelV lHDC, 8, 4, &HD7A483: SetPixelV lHDC, 9, 4, &HDAAC8D: SetPixelV lHDC, 10, 4, &HDBAD8E: SetPixelV lHDC, 11, 4, &HD9AF8C: SetPixelV lHDC, 12, 4, &HDCB28F: SetPixelV lHDC, 13, 4, &HDDB592: SetPixelV lHDC, 14, 4, &HDCB491: SetPixelV lHDC, 15, 4, &HDFB797: SetPixelV lHDC, 16, 4, &HE0B898: SetPixelV lHDC, 17, 4, &HE0B898:
  SetPixelV lHDC, 1, 5, &HFEFEFE: SetPixelV lHDC, 2, 5, &HCDC9C9: SetPixelV lHDC, 3, 5, &H882517: SetPixelV lHDC, 4, 5, &H922100: SetPixelV lHDC, 5, 5, &HA13A00: SetPixelV lHDC, 6, 5, &HD57333: SetPixelV lHDC, 7, 5, &HDFA36F: SetPixelV lHDC, 8, 5, &HDDA876: SetPixelV lHDC, 9, 5, &HD8A573: SetPixelV lHDC, 10, 5, &HDFAE80: SetPixelV lHDC, 11, 5, &HDBAD7D: SetPixelV lHDC, 12, 5, &HDFB084: SetPixelV lHDC, 13, 5, &HDFB286: SetPixelV lHDC, 14, 5, &HDFB188: SetPixelV lHDC, 15, 5, &HE1B58D: SetPixelV lHDC, 16, 5, &HE3B58E: SetPixelV lHDC, 17, 5, &HE3B48E:
  SetPixelV lHDC, 1, 6, &HF9F9F9: SetPixelV lHDC, 2, 6, &H7B706E: SetPixelV lHDC, 3, 6, &H871405: SetPixelV lHDC, 4, 6, &HA5330E: SetPixelV lHDC, 5, 6, &HB34C0D: SetPixelV lHDC, 6, 6, &HD27030: SetPixelV lHDC, 7, 6, &HD89C68: SetPixelV lHDC, 8, 6, &HDAA573: SetPixelV lHDC, 9, 6, &HD9A674: SetPixelV lHDC, 10, 6, &HD9A87A: SetPixelV lHDC, 11, 6, &HDBAD7D: SetPixelV lHDC, 12, 6, &HDBAC80: SetPixelV lHDC, 13, 6, &HDCAF83: SetPixelV lHDC, 14, 6, &HDFB188: SetPixelV lHDC, 15, 6, &HDEB28A: SetPixelV lHDC, 16, 6, &HDFB18A: SetPixelV lHDC, 17, 6, &HE0B18B:
  SetPixelV lHDC, 1, 7, &HE8E8E7: SetPixelV lHDC, 2, 7, &H773F34: SetPixelV lHDC, 3, 7, &H9F2C00: SetPixelV lHDC, 4, 7, &HBA4B07: SetPixelV lHDC, 5, 7, &HC35E10: SetPixelV lHDC, 6, 7, &HCC7323: SetPixelV lHDC, 7, 7, &HDB8F46: SetPixelV lHDC, 8, 7, &HE8A763: SetPixelV lHDC, 9, 7, &HE3A76C: SetPixelV lHDC, 10, 7, &HE7AB70: SetPixelV lHDC, 11, 7, &HE8AE73: SetPixelV lHDC, 12, 7, &HE8AE73: SetPixelV lHDC, 13, 7, &HEDB17B: SetPixelV lHDC, 14, 7, &HEFB37D: SetPixelV lHDC, 15, 7, &HE9B57E: SetPixelV lHDC, 16, 7, &HE9B57E: SetPixelV lHDC, 17, 7, &HE9B47F:
  SetPixelV lHDC, 0, 8, &HFDFDFD: SetPixelV lHDC, 1, 8, &HCAC5C5: SetPixelV lHDC, 2, 8, &H682A1F: SetPixelV lHDC, 3, 8, &HB23E0C: SetPixelV lHDC, 4, 8, &HCC5D19: SetPixelV lHDC, 5, 8, &HCE691B: SetPixelV lHDC, 6, 8, &HCE7525: SetPixelV lHDC, 7, 8, &HCD8138: SetPixelV lHDC, 8, 8, &HC58440: SetPixelV lHDC, 9, 8, &HC5894E: SetPixelV lHDC, 10, 8, &HC98D52: SetPixelV lHDC, 11, 8, &HC88E53: SetPixelV lHDC, 12, 8, &HCC9257: SetPixelV lHDC, 13, 8, &HCF935D: SetPixelV lHDC, 14, 8, &HD0945E: SetPixelV lHDC, 15, 8, &HCE9963: SetPixelV lHDC, 16, 8, &HCE9963: SetPixelV lHDC, 17, 8, &HCE9963:
  SetPixelV lHDC, 0, 9, &HFAFAFA: SetPixelV lHDC, 1, 9, &HB9ADAB: SetPixelV lHDC, 2, 9, &H6E2B10: SetPixelV lHDC, 3, 9, &HB6580D: SetPixelV lHDC, 4, 9, &HCA6C20: SetPixelV lHDC, 5, 9, &HCE792B: SetPixelV lHDC, 6, 9, &HCE8132: SetPixelV lHDC, 7, 9, &HD08B42: SetPixelV lHDC, 8, 9, &HD3904B: SetPixelV lHDC, 9, 9, &HD3934C: SetPixelV lHDC, 10, 9, &HD89753: SetPixelV lHDC, 11, 9, &HDB9B5A: SetPixelV lHDC, 12, 9, &HDC9B5E: SetPixelV lHDC, 13, 9, &HDB9C60: SetPixelV lHDC, 14, 9, &HDB9C60: SetPixelV lHDC, 15, 9, &HDDA164: SetPixelV lHDC, 16, 9, &HDDA164: SetPixelV lHDC, 17, 9, &HDDA064:
  SetPixelV lHDC, 0, 10, &HF7F7F7: SetPixelV lHDC, 1, 10, &HB0A09E: SetPixelV lHDC, 2, 10, &H712E13: SetPixelV lHDC, 3, 10, &HBD5F14: SetPixelV lHDC, 4, 10, &HD17327: SetPixelV lHDC, 5, 10, &HD47F31: SetPixelV lHDC, 6, 10, &HD98C3D: SetPixelV lHDC, 7, 10, &HD9944B: SetPixelV lHDC, 8, 10, &HD7944F: SetPixelV lHDC, 9, 10, &HDC9C55: SetPixelV lHDC, 10, 10, &HDC9B57: SetPixelV lHDC, 11, 10, &HE3A362: SetPixelV lHDC, 12, 10, &HE3A265: SetPixelV lHDC, 13, 10, &HE2A367: SetPixelV lHDC, 14, 10, &HE0A165: SetPixelV lHDC, 15, 10, &HE3A66A: SetPixelV lHDC, 16, 10, &HE3A66A: SetPixelV lHDC, 17, 10, &HE2A66A
  tmph = lh - 22
  SetPixelV lHDC, 0, tmph + 10, &HF7F7F7: SetPixelV lHDC, 1, tmph + 10, &HB0A09E: SetPixelV lHDC, 2, tmph + 10, &H712E13: SetPixelV lHDC, 3, tmph + 10, &HBD5F14: SetPixelV lHDC, 4, tmph + 10, &HD17327: SetPixelV lHDC, 5, tmph + 10, &HD47F31: SetPixelV lHDC, 6, tmph + 10, &HD98C3D: SetPixelV lHDC, 7, tmph + 10, &HD9944B: SetPixelV lHDC, 8, tmph + 10, &HD7944F: SetPixelV lHDC, 9, tmph + 10, &HDC9C55: SetPixelV lHDC, 10, tmph + 10, &HDC9B57: SetPixelV lHDC, 11, tmph + 10, &HE3A362: SetPixelV lHDC, 12, tmph + 10, &HE3A265: SetPixelV lHDC, 13, tmph + 10, &HE2A367: SetPixelV lHDC, 14, tmph + 10, &HE0A165: SetPixelV lHDC, 15, tmph + 10, &HE3A66A: SetPixelV lHDC, 16, tmph + 10, &HE3A66A: SetPixelV lHDC, 17, tmph + 10, &HE2A66A:
  SetPixelV lHDC, 0, tmph + 11, &HF5F5F5: SetPixelV lHDC, 1, tmph + 11, &HACA39E: SetPixelV lHDC, 2, tmph + 11, &H744421: SetPixelV lHDC, 3, tmph + 11, &HC56F1F: SetPixelV lHDC, 4, tmph + 11, &HD17A2A: SetPixelV lHDC, 5, tmph + 11, &HD58C42: SetPixelV lHDC, 6, tmph + 11, &HD7914B: SetPixelV lHDC, 7, tmph + 11, &HDF9854: SetPixelV lHDC, 8, tmph + 11, &HE4A05F: SetPixelV lHDC, 9, tmph + 11, &HE29F66: SetPixelV lHDC, 10, tmph + 11, &HE4A56B: SetPixelV lHDC, 11, tmph + 11, &HDDA467: SetPixelV lHDC, 12, tmph + 11, &HE0A76A: SetPixelV lHDC, 13, tmph + 11, &HE2A96C: SetPixelV lHDC, 14, tmph + 11, &HE3A870: SetPixelV lHDC, 15, tmph + 11, &HE6AC76: SetPixelV lHDC, 16, tmph + 11, &HE6AC76: SetPixelV lHDC, 17, tmph + 11, &HE6AC76:
  SetPixelV lHDC, 0, tmph + 12, &HF5F5F5: SetPixelV lHDC, 1, tmph + 12, &HB1AAA7: SetPixelV lHDC, 2, tmph + 12, &H825533: SetPixelV lHDC, 3, tmph + 12, &HCF792A: SetPixelV lHDC, 4, tmph + 12, &HE48D3D: SetPixelV lHDC, 5, tmph + 12, &HDD944A: SetPixelV lHDC, 6, tmph + 12, &HE49E58: SetPixelV lHDC, 7, tmph + 12, &HEBA460: SetPixelV lHDC, 8, tmph + 12, &HEEAA69: SetPixelV lHDC, 9, tmph + 12, &HF3B077: SetPixelV lHDC, 10, tmph + 12, &HEEAF75: SetPixelV lHDC, 11, tmph + 12, &HEBB275: SetPixelV lHDC, 12, tmph + 12, &HEFB679: SetPixelV lHDC, 13, tmph + 12, &HF1B87B: SetPixelV lHDC, 14, tmph + 12, &HF1B67E: SetPixelV lHDC, 15, tmph + 12, &HF2B781: SetPixelV lHDC, 16, tmph + 12, &HF1B681: SetPixelV lHDC, 17, tmph + 12, &HF1B681:
  SetPixelV lHDC, 0, tmph + 13, &HF7F7F7: SetPixelV lHDC, 1, tmph + 13, &HC2C2C1: SetPixelV lHDC, 2, tmph + 13, &H6B5D4E: SetPixelV lHDC, 3, tmph + 13, &HC27831: SetPixelV lHDC, 4, tmph + 13, &HDA8E46: SetPixelV lHDC, 5, tmph + 13, &HE7A05C: SetPixelV lHDC, 6, tmph + 13, &HEAA665: SetPixelV lHDC, 7, tmph + 13, &HE9AF6E: SetPixelV lHDC, 8, tmph + 13, &HEFB377: SetPixelV lHDC, 9, tmph + 13, &HF3B579: SetPixelV lHDC, 10, tmph + 13, &HF7B97D: SetPixelV lHDC, 11, tmph + 13, &HF2BB7E: SetPixelV lHDC, 12, tmph + 13, &HF4BB83: SetPixelV lHDC, 13, tmph + 13, &HF5BE85: SetPixelV lHDC, 14, tmph + 13, &HF4BB87: SetPixelV lHDC, 15, tmph + 13, &HF5BE8A: SetPixelV lHDC, 16, tmph + 13, &HF5BD8A: SetPixelV lHDC, 17, tmph + 13, &HF3BD8A:
  SetPixelV lHDC, 0, tmph + 14, &HFBFBFB: SetPixelV lHDC, 1, tmph + 14, &HE1E1E1: SetPixelV lHDC, 2, tmph + 14, &H85796E: SetPixelV lHDC, 3, tmph + 14, &HB76F2B: SetPixelV lHDC, 4, tmph + 14, &HDE924A: SetPixelV lHDC, 5, tmph + 14, &HE8A15D: SetPixelV lHDC, 6, tmph + 14, &HF2AE6D: SetPixelV lHDC, 7, tmph + 14, &HF1B776: SetPixelV lHDC, 8, tmph + 14, &HF2B67A: SetPixelV lHDC, 9, tmph + 14, &HFBBD81: SetPixelV lHDC, 10, tmph + 14, &HFFC286: SetPixelV lHDC, 11, tmph + 14, &HFAC386: SetPixelV lHDC, 12, tmph + 14, &HFBC28A: SetPixelV lHDC, 13, tmph + 14, &HFAC38A: SetPixelV lHDC, 14, tmph + 14, &HFAC18D: SetPixelV lHDC, 15, tmph + 14, &HFDC592: SetPixelV lHDC, 16, tmph + 14, &HFDC592: SetPixelV lHDC, 17, tmph + 14, &HFCC592:
  SetPixelV lHDC, 0, tmph + 15, &HFEFEFE: SetPixelV lHDC, 1, tmph + 15, &HEDEDED: SetPixelV lHDC, 2, tmph + 15, &HA2A0A0: SetPixelV lHDC, 3, tmph + 15, &H816753: SetPixelV lHDC, 4, tmph + 15, &HC09068: SetPixelV lHDC, 5, tmph + 15, &HEDA55F: SetPixelV lHDC, 6, tmph + 15, &HFAB26C: SetPixelV lHDC, 7, tmph + 15, &HFCBF7D: SetPixelV lHDC, 8, tmph + 15, &HF7C182: SetPixelV lHDC, 9, tmph + 15, &HF8C38A: SetPixelV lHDC, 10, tmph + 15, &HFACA90: SetPixelV lHDC, 11, tmph + 15, &HF7CB8E: SetPixelV lHDC, 12, tmph + 15, &HF8CC8F: SetPixelV lHDC, 13, tmph + 15, &HFACC96: SetPixelV lHDC, 14, tmph + 15, &HF9CB95: SetPixelV lHDC, 15, tmph + 15, &HF9CE97: SetPixelV lHDC, 16, tmph + 15, &HF8CD97: SetPixelV lHDC, 17, tmph + 15, &HF8CE97:
  SetPixelV lHDC, 1, tmph + 16, &HF6F6F6: SetPixelV lHDC, 2, tmph + 16, &HD6D6D6: SetPixelV lHDC, 3, tmph + 16, &H8E7C6F: SetPixelV lHDC, 4, tmph + 16, &H946843: SetPixelV lHDC, 5, tmph + 16, &HEEA762: SetPixelV lHDC, 6, tmph + 16, &HFFB771: SetPixelV lHDC, 7, tmph + 16, &HFEC17F: SetPixelV lHDC, 8, tmph + 16, &HFFC98A: SetPixelV lHDC, 9, tmph + 16, &HFFCE95: SetPixelV lHDC, 10, tmph + 16, &HFBCB91: SetPixelV lHDC, 11, tmph + 16, &HFFD396: SetPixelV lHDC, 12, tmph + 16, &HFFD396: SetPixelV lHDC, 13, tmph + 16, &HFFD29C: SetPixelV lHDC, 14, tmph + 16, &HFFD39D: SetPixelV lHDC, 15, tmph + 16, &HFFD49E: SetPixelV lHDC, 16, tmph + 16, &HFFD49E: SetPixelV lHDC, 17, tmph + 16, &HFED59E:
  SetPixelV lHDC, 1, tmph + 17, &HFDFDFD: SetPixelV lHDC, 2, tmph + 17, &HEDEDED: SetPixelV lHDC, 3, tmph + 17, &HBEBEBE: SetPixelV lHDC, 4, tmph + 17, &H6C6C6C: SetPixelV lHDC, 5, tmph + 17, &H7C684F: SetPixelV lHDC, 6, tmph + 17, &HD1AE81: SetPixelV lHDC, 7, tmph + 17, &HF1C284: SetPixelV lHDC, 8, tmph + 17, &HFDCE90: SetPixelV lHDC, 9, tmph + 17, &HF8D193: SetPixelV lHDC, 10, tmph + 17, &HFBD899: SetPixelV lHDC, 11, tmph + 17, &HF5DC9E: SetPixelV lHDC, 12, tmph + 17, &HF8DFA1: SetPixelV lHDC, 13, tmph + 17, &HF8DFA1: SetPixelV lHDC, 14, tmph + 17, &HF8DFA1: SetPixelV lHDC, 15, tmph + 17, &HF8DEA3: SetPixelV lHDC, 16, tmph + 17, &HF7DDA3: SetPixelV lHDC, 17, tmph + 17, &HF7DDA3:
  SetPixelV lHDC, 2, tmph + 18, &HF9F9F9: SetPixelV lHDC, 3, tmph + 18, &HE6E6E6: SetPixelV lHDC, 4, tmph + 18, &HBABABA: SetPixelV lHDC, 5, tmph + 18, &H827666: SetPixelV lHDC, 6, tmph + 18, &H836743: SetPixelV lHDC, 7, tmph + 18, &HBE935B: SetPixelV lHDC, 8, tmph + 18, &HF4C78B: SetPixelV lHDC, 9, tmph + 18, &HFDD79A: SetPixelV lHDC, 10, tmph + 18, &HFFDFA0: SetPixelV lHDC, 11, tmph + 18, &HFBE2A4: SetPixelV lHDC, 12, tmph + 18, &HFFE7A9: SetPixelV lHDC, 13, tmph + 18, &HFFE9AB: SetPixelV lHDC, 14, tmph + 18, &HFFE7A9: SetPixelV lHDC, 15, tmph + 18, &HFFE6AC: SetPixelV lHDC, 16, tmph + 18, &HFFE6AD: SetPixelV lHDC, 17, tmph + 18, &HFFE6AD:
  SetPixelV lHDC, 2, tmph + 19, &HFEFEFE: SetPixelV lHDC, 3, tmph + 19, &HF8F8F8: SetPixelV lHDC, 4, tmph + 19, &HE6E6E6: SetPixelV lHDC, 5, tmph + 19, &HC8C8C8: SetPixelV lHDC, 6, tmph + 19, &H8F8F8F: SetPixelV lHDC, 7, tmph + 19, &H686462: SetPixelV lHDC, 8, tmph + 19, &H6D655E: SetPixelV lHDC, 9, tmph + 19, &H918472: SetPixelV lHDC, 10, tmph + 19, &HB3A88E: SetPixelV lHDC, 11, tmph + 19, &HDAD1B2: SetPixelV lHDC, 12, tmph + 19, &HE3DBBA: SetPixelV lHDC, 13, tmph + 19, &HE7E0C0: SetPixelV lHDC, 14, tmph + 19, &HE9E2C1: SetPixelV lHDC, 15, tmph + 19, &HE9E2C5: SetPixelV lHDC, 16, tmph + 19, &HE9E1C5: SetPixelV lHDC, 17, tmph + 19, &HE9E2C5:
  SetPixelV lHDC, 3, tmph + 20, &HFEFEFE: SetPixelV lHDC, 4, tmph + 20, &HF9F9F9: SetPixelV lHDC, 5, tmph + 20, &HECECEC: SetPixelV lHDC, 6, tmph + 20, &HDADADA: SetPixelV lHDC, 7, tmph + 20, &HC2C2C1: SetPixelV lHDC, 8, tmph + 20, &H9F9D9B: SetPixelV lHDC, 9, tmph + 20, &H827D75: SetPixelV lHDC, 10, tmph + 20, &H6A6353: SetPixelV lHDC, 11, tmph + 20, &H5F5941: SetPixelV lHDC, 12, tmph + 20, &H5D553B: SetPixelV lHDC, 13, tmph + 20, &H595338: SetPixelV lHDC, 14, tmph + 20, &H5E5739: SetPixelV lHDC, 15, tmph + 20, &H5F5A3C: SetPixelV lHDC, 16, tmph + 20, &H635E3F: SetPixelV lHDC, 17, tmph + 20, &H635D40:
  SetPixelV lHDC, 5, tmph + 21, &HFCFCFC: SetPixelV lHDC, 6, tmph + 21, &HF5F5F5: SetPixelV lHDC, 7, tmph + 21, &HEBEBEB: SetPixelV lHDC, 8, tmph + 21, &HE1E1E1: SetPixelV lHDC, 9, tmph + 21, &HD6D6D6: SetPixelV lHDC, 10, tmph + 21, &HCECECE: SetPixelV lHDC, 11, tmph + 21, &HC9C9C9: SetPixelV lHDC, 12, tmph + 21, &HC7C7C7: SetPixelV lHDC, 13, tmph + 21, &HC7C7C7: SetPixelV lHDC, 14, tmph + 21, &HC6C6C6: SetPixelV lHDC, 15, tmph + 21, &HC6C6C6: SetPixelV lHDC, 16, tmph + 21, &HC5C5C5: SetPixelV lHDC, 17, tmph + 21, &HC5C5C5:
  SetPixelV lHDC, 7, tmph + 22, &HFDFDFD: SetPixelV lHDC, 8, tmph + 22, &HF9F9F9: SetPixelV lHDC, 9, tmph + 22, &HF4F4F4: SetPixelV lHDC, 10, tmph + 22, &HF0F0F0: SetPixelV lHDC, 11, tmph + 22, &HEEEEEE: SetPixelV lHDC, 12, tmph + 22, &HEDEDED: SetPixelV lHDC, 13, tmph + 22, &HECECEC: SetPixelV lHDC, 14, tmph + 22, &HECECEC: SetPixelV lHDC, 15, tmph + 22, &HECECEC: SetPixelV lHDC, 16, tmph + 22, &HECECEC: SetPixelV lHDC, 17, tmph + 22, &HECECEC:
  tmpw = lw - 34
  SetPixelV lHDC, tmpw + 17, 0, &H450608: SetPixelV lHDC, tmpw + 18, 0, &H450608: SetPixelV lHDC, tmpw + 19, 0, &H3B0707: SetPixelV lHDC, tmpw + 20, 0, &H370706: SetPixelV lHDC, tmpw + 21, 0, &H360507: SetPixelV lHDC, tmpw + 22, 0, &H3B0F10: SetPixelV lHDC, tmpw + 23, 0, &H442526: SetPixelV lHDC, tmpw + 24, 0, &H604E4E: SetPixelV lHDC, tmpw + 25, 0, &HA29D9E: SetPixelV lHDC, tmpw + 26, 0, &HEEEEEE: SetPixelV lHDC, tmpw + 34, 0, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 1, &HF1D4C9: SetPixelV lHDC, tmpw + 18, 1, &HF1D4C9: SetPixelV lHDC, tmpw + 19, 1, &HEDD3CD: SetPixelV lHDC, tmpw + 20, 1, &HEBD1CB: SetPixelV lHDC, tmpw + 21, 1, &HE9CEC4: SetPixelV lHDC, tmpw + 22, 1, &HE5C1B9: SetPixelV lHDC, tmpw + 23, 1, &HCFA89F: SetPixelV lHDC, tmpw + 24, 1, &HAA6E68: SetPixelV lHDC, tmpw + 25, 1, &H73211B: SetPixelV lHDC, tmpw + 26, 1, &H702924: SetPixelV lHDC, tmpw + 27, 1, &HAA9897: SetPixelV lHDC, tmpw + 28, 1, &HF7F7F7: SetPixelV lHDC, tmpw + 34, 1, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 2, &HE5C8BD: SetPixelV lHDC, tmpw + 18, 2, &HE5C8BD: SetPixelV lHDC, tmpw + 19, 2, &HDEC4BE: SetPixelV lHDC, tmpw + 20, 2, &HDCC2BC: SetPixelV lHDC, tmpw + 21, 2, &HE2C7BD: SetPixelV lHDC, tmpw + 22, 2, &HE2BEB6: SetPixelV lHDC, tmpw + 23, 2, &HE8C1B8: SetPixelV lHDC, tmpw + 24, 2, &HF0B4AE: SetPixelV lHDC, tmpw + 25, 2, &HF29C96: SetPixelV lHDC, tmpw + 26, 2, &H822D27: SetPixelV lHDC, tmpw + 27, 2, &H400807: SetPixelV lHDC, tmpw + 28, 2, &H71585A: SetPixelV lHDC, tmpw + 29, 2, &HEBEBEB: SetPixelV lHDC, tmpw + 34, 2, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 3, &HE8C0A1: SetPixelV lHDC, tmpw + 18, 3, &HE8C0A1: SetPixelV lHDC, tmpw + 19, 3, &HE5C09A: SetPixelV lHDC, tmpw + 20, 3, &HE4BF99: SetPixelV lHDC, tmpw + 21, 3, &HE4BA97: SetPixelV lHDC, tmpw + 22, 3, &HE9BF9C: SetPixelV lHDC, tmpw + 23, 3, &HDFB695: SetPixelV lHDC, tmpw + 24, 3, &HDFB695: SetPixelV lHDC, tmpw + 25, 3, &HE0AE90: SetPixelV lHDC, tmpw + 26, 3, &HCB8469: SetPixelV lHDC, tmpw + 27, 3, &H941600: SetPixelV lHDC, tmpw + 28, 3, &H830800: SetPixelV lHDC, tmpw + 29, 3, &H895253: SetPixelV lHDC, tmpw + 30, 3, &HF0EFEF: SetPixelV lHDC, tmpw + 34, 3, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 4, &HE0B898: SetPixelV lHDC, tmpw + 18, 4, &HE0B897: SetPixelV lHDC, tmpw + 19, 4, &HDAB58F: SetPixelV lHDC, tmpw + 20, 4, &HDBB690: SetPixelV lHDC, tmpw + 21, 4, &HDBB18E: SetPixelV lHDC, tmpw + 22, 4, &HD7AD8A: SetPixelV lHDC, tmpw + 23, 4, &HDAB190: SetPixelV lHDC, tmpw + 24, 4, &HD2A988: SetPixelV lHDC, tmpw + 25, 4, &HD6A486: SetPixelV lHDC, tmpw + 26, 4, &HDA9378: SetPixelV lHDC, tmpw + 27, 4, &HBF4129: SetPixelV lHDC, tmpw + 28, 4, &H991B03: SetPixelV lHDC, tmpw + 29, 4, &H500709: SetPixelV lHDC, tmpw + 30, 4, &H826F70: SetPixelV lHDC, tmpw + 31, 4, &HFAFAFA: SetPixelV lHDC, tmpw + 34, 4, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 5, &HE3B48E: SetPixelV lHDC, tmpw + 18, 5, &HE3B48D: SetPixelV lHDC, tmpw + 19, 5, &HE0B387: SetPixelV lHDC, tmpw + 20, 5, &HDEB185: SetPixelV lHDC, tmpw + 21, 5, &HE1B084: SetPixelV lHDC, tmpw + 22, 5, &HE3AE83: SetPixelV lHDC, tmpw + 23, 5, &HE1AF7B: SetPixelV lHDC, tmpw + 24, 5, &HE0A976: SetPixelV lHDC, tmpw + 25, 5, &HDCA473: SetPixelV lHDC, tmpw + 26, 5, &HDEA372: SetPixelV lHDC, tmpw + 27, 5, &HCC712E: SetPixelV lHDC, tmpw + 28, 5, &HA53900: SetPixelV lHDC, tmpw + 29, 5, &H9D2200: SetPixelV lHDC, tmpw + 30, 5, &H9E2114: SetPixelV lHDC, tmpw + 31, 5, &HC7C5C4: SetPixelV lHDC, tmpw + 32, 5, &HFEFEFE: SetPixelV lHDC, tmpw + 34, 5, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 6, &HE0B18B: SetPixelV lHDC, tmpw + 18, 6, &HE0B18A: SetPixelV lHDC, tmpw + 19, 6, &HDEB185: SetPixelV lHDC, tmpw + 20, 6, &HDEB185: SetPixelV lHDC, tmpw + 21, 6, &HDCAB7F: SetPixelV lHDC, tmpw + 22, 6, &HE1AC81: SetPixelV lHDC, tmpw + 23, 6, &HDCAA76: SetPixelV lHDC, tmpw + 24, 6, &HDCA572: SetPixelV lHDC, tmpw + 25, 6, &HDBA372: SetPixelV lHDC, tmpw + 26, 6, &HD79C6B: SetPixelV lHDC, tmpw + 27, 6, &HD17633: SetPixelV lHDC, tmpw + 28, 6, &HB74B0B: SetPixelV lHDC, tmpw + 29, 6, &HAC310D: SetPixelV lHDC, tmpw + 30, 6, &H961507: SetPixelV lHDC, tmpw + 31, 6, &H736D6A: SetPixelV lHDC, tmpw + 32, 6, &HF8F8F8: SetPixelV lHDC, tmpw + 34, 6, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 7, &HE9B47F: SetPixelV lHDC, tmpw + 18, 7, &HEAB47E: SetPixelV lHDC, tmpw + 19, 7, &HEFB67E: SetPixelV lHDC, tmpw + 20, 7, &HE8AF77: SetPixelV lHDC, tmpw + 21, 7, &HE7AF74: SetPixelV lHDC, tmpw + 22, 7, &HE4AC71: SetPixelV lHDC, tmpw + 23, 7, &HEAAD6F: SetPixelV lHDC, tmpw + 24, 7, &HE9A968: SetPixelV lHDC, tmpw + 25, 7, &HE7A564: SetPixelV lHDC, tmpw + 26, 7, &HD9904C: SetPixelV lHDC, tmpw + 27, 7, &HC5711F: SetPixelV lHDC, tmpw + 28, 7, &HC16010: SetPixelV lHDC, tmpw + 29, 7, &HBB4D05: SetPixelV lHDC, tmpw + 30, 7, &HA02D00: SetPixelV lHDC, tmpw + 31, 7, &H774033: SetPixelV lHDC, tmpw + 32, 7, &HE7E6E6: SetPixelV lHDC, tmpw + 34, 7, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 8, &HCE9963: SetPixelV lHDC, tmpw + 18, 8, &HCF9963: SetPixelV lHDC, tmpw + 19, 8, &HCE955D: SetPixelV lHDC, tmpw + 20, 8, &HCE955D: SetPixelV lHDC, tmpw + 21, 8, &HCA9257: SetPixelV lHDC, tmpw + 22, 8, &HC89055: SetPixelV lHDC, tmpw + 23, 8, &HCB8E50: SetPixelV lHDC, tmpw + 24, 8, &HCB8B4A: SetPixelV lHDC, tmpw + 25, 8, &HC58342: SetPixelV lHDC, tmpw + 26, 8, &HC87F3B: SetPixelV lHDC, tmpw + 27, 8, &HCA7624: SetPixelV lHDC, tmpw + 28, 8, &HCA6919: SetPixelV lHDC, tmpw + 29, 8, &HCC5E16: SetPixelV lHDC, tmpw + 30, 8, &HB23E07: SetPixelV lHDC, tmpw + 31, 8, &H682B1D: SetPixelV lHDC, tmpw + 32, 8, &HC7C2C2: SetPixelV lHDC, tmpw + 33, 8, &HFDFDFD: SetPixelV lHDC, tmpw + 34, 8, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 9, &HDDA064: SetPixelV lHDC, tmpw + 18, 9, &HDCA064: SetPixelV lHDC, tmpw + 19, 9, &HDA9D5D: SetPixelV lHDC, tmpw + 20, 9, &HD99C5C: SetPixelV lHDC, tmpw + 21, 9, &HDA9D5D: SetPixelV lHDC, tmpw + 22, 9, &HDA9A5A: SetPixelV lHDC, tmpw + 23, 9, &HD89753: SetPixelV lHDC, tmpw + 24, 9, &HD7914E: SetPixelV lHDC, tmpw + 25, 9, &HD38E49: SetPixelV lHDC, tmpw + 26, 9, &HD38B43: SetPixelV lHDC, tmpw + 27, 9, &HCD8430: SetPixelV lHDC, tmpw + 28, 9, &HCA7826: SetPixelV lHDC, tmpw + 29, 9, &HCE6C1E: SetPixelV lHDC, tmpw + 30, 9, &HB9560C: SetPixelV lHDC, tmpw + 31, 9, &H742E0D: SetPixelV lHDC, tmpw + 32, 9, &HB3A6A4: SetPixelV lHDC, tmpw + 33, 9, &HF9F9F9: SetPixelV lHDC, tmpw + 34, 9, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 10, &HE2A66A: SetPixelV lHDC, tmpw + 18, 10, &HE2A66A: SetPixelV lHDC, tmpw + 19, 10, &HE1A464: SetPixelV lHDC, tmpw + 20, 10, &HE0A363: SetPixelV lHDC, tmpw + 21, 10, &HE0A363: SetPixelV lHDC, tmpw + 22, 10, &HE1A161: SetPixelV lHDC, tmpw + 23, 10, &HE09F5B: SetPixelV lHDC, tmpw + 24, 10, &HDE9855: SetPixelV lHDC, tmpw + 25, 10, &HDC9752: SetPixelV lHDC, tmpw + 26, 10, &HDB934B: SetPixelV lHDC, tmpw + 27, 10, &HD68D39: SetPixelV lHDC, tmpw + 28, 10, &HD17F2D: SetPixelV lHDC, tmpw + 29, 10, &HD67426: SetPixelV lHDC, tmpw + 30, 10, &HC05D13: SetPixelV lHDC, tmpw + 31, 10, &H7C3514: SetPixelV lHDC, tmpw + 32, 10, &HAB9B98: SetPixelV lHDC, tmpw + 33, 10, &HF6F6F6: SetPixelV lHDC, tmpw + 34, 10, &HFFFFFFFF:
  tmph = lh - 22
  tmpw = lw - 34
  SetPixelV lHDC, tmpw + 17, tmph + 10, &HE2A66A: SetPixelV lHDC, tmpw + 18, tmph + 10, &HE2A66A: SetPixelV lHDC, tmpw + 19, tmph + 10, &HE1A464: SetPixelV lHDC, tmpw + 20, tmph + 10, &HE0A363: SetPixelV lHDC, tmpw + 21, tmph + 10, &HE0A363: SetPixelV lHDC, tmpw + 22, tmph + 10, &HE1A161: SetPixelV lHDC, tmpw + 23, tmph + 10, &HE09F5B: SetPixelV lHDC, tmpw + 24, tmph + 10, &HDE9855: SetPixelV lHDC, tmpw + 25, tmph + 10, &HDC9752: SetPixelV lHDC, tmpw + 26, tmph + 10, &HDB934B: SetPixelV lHDC, tmpw + 27, tmph + 10, &HD68D39: SetPixelV lHDC, tmpw + 28, tmph + 10, &HD17F2D: SetPixelV lHDC, tmpw + 29, tmph + 10, &HD67426: SetPixelV lHDC, tmpw + 30, tmph + 10, &HC05D13: SetPixelV lHDC, tmpw + 31, tmph + 10, &H7C3514: SetPixelV lHDC, tmpw + 32, tmph + 10, &HAB9B98: SetPixelV lHDC, tmpw + 33, tmph + 10, &HF6F6F6:
  SetPixelV lHDC, tmpw + 17, tmph + 11, &HE6AC76: SetPixelV lHDC, tmpw + 18, tmph + 11, &HE6AC76: SetPixelV lHDC, tmpw + 19, tmph + 11, &HE2A86D: SetPixelV lHDC, tmpw + 20, tmph + 11, &HE5A66C: SetPixelV lHDC, tmpw + 21, tmph + 11, &HE1A56A: SetPixelV lHDC, tmpw + 22, tmph + 11, &HE4A46A: SetPixelV lHDC, tmpw + 23, tmph + 11, &HE1A266: SetPixelV lHDC, tmpw + 24, tmph + 11, &HE6A364: SetPixelV lHDC, tmpw + 25, tmph + 11, &HE19F5E: SetPixelV lHDC, tmpw + 26, tmph + 11, &HDF9A55: SetPixelV lHDC, tmpw + 27, tmph + 11, &HD89048: SetPixelV lHDC, tmpw + 28, tmph + 11, &HD88A3E: SetPixelV lHDC, tmpw + 29, tmph + 11, &HCF7927: SetPixelV lHDC, tmpw + 30, tmph + 11, &HC87220: SetPixelV lHDC, tmpw + 31, tmph + 11, &H77481E: SetPixelV lHDC, tmpw + 32, tmph + 11, &HABA39E: SetPixelV lHDC, tmpw + 33, tmph + 11, &HF4F4F4:
  SetPixelV lHDC, tmpw + 17, tmph + 12, &HF1B681: SetPixelV lHDC, tmpw + 18, tmph + 12, &HF0B780: SetPixelV lHDC, tmpw + 19, tmph + 12, &HF2B87D: SetPixelV lHDC, tmpw + 20, tmph + 12, &HF5B67C: SetPixelV lHDC, tmpw + 21, tmph + 12, &HF1B57A: SetPixelV lHDC, tmpw + 22, tmph + 12, &HF2B278: SetPixelV lHDC, tmpw + 23, tmph + 12, &HF0B175: SetPixelV lHDC, tmpw + 24, tmph + 12, &HF3B071: SetPixelV lHDC, tmpw + 25, tmph + 12, &HECAA69: SetPixelV lHDC, tmpw + 26, tmph + 12, &HE9A45F: SetPixelV lHDC, tmpw + 27, tmph + 12, &HE8A058: SetPixelV lHDC, tmpw + 28, tmph + 12, &HE5974B: SetPixelV lHDC, tmpw + 29, tmph + 12, &HE38D3B: SetPixelV lHDC, tmpw + 30, tmph + 12, &HD37D2B: SetPixelV lHDC, tmpw + 31, tmph + 12, &H895A32: SetPixelV lHDC, tmpw + 32, tmph + 12, &HB4AFAC: SetPixelV lHDC, tmpw + 33, tmph + 12, &HF5F5F5:
  SetPixelV lHDC, tmpw + 17, tmph + 13, &HF3BD8A: SetPixelV lHDC, tmpw + 18, tmph + 13, &HF3BD8A: SetPixelV lHDC, tmpw + 19, tmph + 13, &HF2BD84: SetPixelV lHDC, tmpw + 20, tmph + 13, &HF5BC84: SetPixelV lHDC, tmpw + 21, tmph + 13, &HF3BC83: SetPixelV lHDC, tmpw + 22, tmph + 13, &HF4B981: SetPixelV lHDC, tmpw + 23, tmph + 13, &HF2B97C: SetPixelV lHDC, tmpw + 24, tmph + 13, &HF5B77B: SetPixelV lHDC, tmpw + 25, tmph + 13, &HF1B476: SetPixelV lHDC, tmpw + 26, tmph + 13, &HEFAF6E: SetPixelV lHDC, tmpw + 27, tmph + 13, &HE5A45F: SetPixelV lHDC, tmpw + 28, tmph + 13, &HE49F5A: SetPixelV lHDC, tmpw + 29, tmph + 13, &HDA8F4A: SetPixelV lHDC, tmpw + 30, tmph + 13, &HC57A35: SetPixelV lHDC, tmpw + 31, tmph + 13, &H736353: SetPixelV lHDC, tmpw + 32, tmph + 13, &HD6D5D5: SetPixelV lHDC, tmpw + 33, tmph + 13, &HF8F8F8:
  SetPixelV lHDC, tmpw + 17, tmph + 14, &HFCC592: SetPixelV lHDC, tmpw + 18, tmph + 14, &HFBC592: SetPixelV lHDC, tmpw + 19, tmph + 14, &HF7C289: SetPixelV lHDC, tmpw + 20, tmph + 14, &HFCC38B: SetPixelV lHDC, tmpw + 21, tmph + 14, &HFAC38A: SetPixelV lHDC, tmpw + 22, tmph + 14, &HFDC28A: SetPixelV lHDC, tmpw + 23, tmph + 14, &HFBC285: SetPixelV lHDC, tmpw + 24, tmph + 14, &HFBBD81: SetPixelV lHDC, tmpw + 25, tmph + 14, &HF6B97B: SetPixelV lHDC, tmpw + 26, tmph + 14, &HF6B675: SetPixelV lHDC, tmpw + 27, tmph + 14, &HF0AF6A: SetPixelV lHDC, tmpw + 28, tmph + 14, &HE8A35E: SetPixelV lHDC, tmpw + 29, tmph + 14, &HDD924D: SetPixelV lHDC, tmpw + 30, tmph + 14, &HBA702B: SetPixelV lHDC, tmpw + 31, tmph + 14, &H847A70: SetPixelV lHDC, tmpw + 32, tmph + 14, &HE8E8E8: SetPixelV lHDC, tmpw + 33, tmph + 14, &HFDFDFD:
  SetPixelV lHDC, tmpw + 17, tmph + 15, &HF8CE97: SetPixelV lHDC, tmpw + 18, tmph + 15, &HF9CD97: SetPixelV lHDC, tmpw + 19, tmph + 15, &HF9CE95: SetPixelV lHDC, tmpw + 20, tmph + 15, &HF7CC93: SetPixelV lHDC, tmpw + 21, tmph + 15, &HF6CB92: SetPixelV lHDC, tmpw + 22, tmph + 15, &HF9CA92: SetPixelV lHDC, tmpw + 23, tmph + 15, &HFCCD90: SetPixelV lHDC, tmpw + 24, tmph + 15, &HF8C488: SetPixelV lHDC, tmpw + 25, tmph + 15, &HF3BD80: SetPixelV lHDC, tmpw + 26, tmph + 15, &HFABD7D: SetPixelV lHDC, tmpw + 27, tmph + 15, &HF7B26D: SetPixelV lHDC, tmpw + 28, tmph + 15, &HEAA560: SetPixelV lHDC, tmpw + 29, tmph + 15, &HC0925D: SetPixelV lHDC, tmpw + 30, tmph + 15, &H896F54: SetPixelV lHDC, tmpw + 31, tmph + 15, &HBABABB: SetPixelV lHDC, tmpw + 32, tmph + 15, &HF1F1F1:
  SetPixelV lHDC, tmpw + 17, tmph + 16, &HFED59E: SetPixelV lHDC, tmpw + 18, tmph + 16, &HFFD59F: SetPixelV lHDC, tmpw + 19, tmph + 16, &HFED39A: SetPixelV lHDC, tmpw + 20, tmph + 16, &HFFD49B: SetPixelV lHDC, tmpw + 21, tmph + 16, &HFCD198: SetPixelV lHDC, tmpw + 22, tmph + 16, &HFFD098: SetPixelV lHDC, tmpw + 23, tmph + 16, &HFECF92: SetPixelV lHDC, tmpw + 24, tmph + 16, &HFFCB8F: SetPixelV lHDC, tmpw + 25, tmph + 16, &HFFC98C: SetPixelV lHDC, tmpw + 26, tmph + 16, &HFEC181: SetPixelV lHDC, tmpw + 27, tmph + 16, &HFBB671: SetPixelV lHDC, tmpw + 28, tmph + 16, &HF0AB66: SetPixelV lHDC, tmpw + 29, tmph + 16, &H9F733E: SetPixelV lHDC, tmpw + 30, tmph + 16, &H918478: SetPixelV lHDC, tmpw + 31, tmph + 16, &HE2E2E2: SetPixelV lHDC, tmpw + 32, tmph + 16, &HF9F9F9:
  SetPixelV lHDC, tmpw + 17, tmph + 17, &HF7DDA3: SetPixelV lHDC, tmpw + 18, tmph + 17, &HF8DDA4: SetPixelV lHDC, tmpw + 19, tmph + 17, &HF9E0A2: SetPixelV lHDC, tmpw + 20, tmph + 17, &HF5DC9E: SetPixelV lHDC, tmpw + 21, tmph + 17, &HF8DEA2: SetPixelV lHDC, tmpw + 22, tmph + 17, &HFBDDA2: SetPixelV lHDC, tmpw + 23, tmph + 17, &HF7D495: SetPixelV lHDC, tmpw + 24, tmph + 17, &HF8D193: SetPixelV lHDC, tmpw + 25, tmph + 17, &HFCCD90: SetPixelV lHDC, tmpw + 26, tmph + 17, &HF1C088: SetPixelV lHDC, tmpw + 27, tmph + 17, &HDBB186: SetPixelV lHDC, tmpw + 28, tmph + 17, &H8C7259: SetPixelV lHDC, tmpw + 29, tmph + 17, &H6D6B6B: SetPixelV lHDC, tmpw + 30, tmph + 17, &HD2D2D2: SetPixelV lHDC, tmpw + 31, tmph + 17, &HF2F2F2: SetPixelV lHDC, tmpw + 32, tmph + 17, &HFEFEFE:
  SetPixelV lHDC, tmpw + 17, tmph + 18, &HFFE6AD: SetPixelV lHDC, tmpw + 18, tmph + 18, &HFFE6AD: SetPixelV lHDC, tmpw + 19, tmph + 18, &HFFE7A9: SetPixelV lHDC, tmpw + 20, tmph + 18, &HFFEAAC: SetPixelV lHDC, tmpw + 21, tmph + 18, &HF7DDA1: SetPixelV lHDC, tmpw + 22, tmph + 18, &HFFE1A6: SetPixelV lHDC, tmpw + 23, tmph + 18, &HFFE1A2: SetPixelV lHDC, tmpw + 24, tmph + 18, &HFED799: SetPixelV lHDC, tmpw + 25, tmph + 18, &HFACC8F: SetPixelV lHDC, tmpw + 26, tmph + 18, &HC99A64: SetPixelV lHDC, tmpw + 27, tmph + 18, &H977048: SetPixelV lHDC, tmpw + 28, tmph + 18, &H817060: SetPixelV lHDC, tmpw + 29, tmph + 18, &HC8C8C8: SetPixelV lHDC, tmpw + 30, tmph + 18, &HEBEBEB: SetPixelV lHDC, tmpw + 31, tmph + 18, &HFCFCFC:
  SetPixelV lHDC, tmpw + 17, tmph + 19, &HE9E2C5: SetPixelV lHDC, tmpw + 18, tmph + 19, &HE9E2C5: SetPixelV lHDC, tmpw + 19, tmph + 19, &HEAE2C4: SetPixelV lHDC, tmpw + 20, tmph + 19, &HE7DFC1: SetPixelV lHDC, tmpw + 21, tmph + 19, &HEEE4C6: SetPixelV lHDC, tmpw + 22, tmph + 19, &HDBD1B4: SetPixelV lHDC, tmpw + 23, tmph + 19, &HB7AF93: SetPixelV lHDC, tmpw + 24, tmph + 19, &H8D8973: SetPixelV lHDC, tmpw + 25, tmph + 19, &H736D60: SetPixelV lHDC, tmpw + 26, tmph + 19, &H6A6660: SetPixelV lHDC, tmpw + 27, tmph + 19, &H8E9090: SetPixelV lHDC, tmpw + 28, tmph + 19, &HCDCDCD: SetPixelV lHDC, tmpw + 29, tmph + 19, &HE8E8E8: SetPixelV lHDC, tmpw + 30, tmph + 19, &HFAFAFA:
  SetPixelV lHDC, tmpw + 17, tmph + 20, &H635D40: SetPixelV lHDC, tmpw + 18, tmph + 20, &H615B3F: SetPixelV lHDC, tmpw + 19, tmph + 20, &H60583C: SetPixelV lHDC, tmpw + 20, tmph + 20, &H5D563A: SetPixelV lHDC, tmpw + 21, tmph + 20, &H61583D: SetPixelV lHDC, tmpw + 22, tmph + 20, &H605840: SetPixelV lHDC, tmpw + 23, tmph + 20, &H6A6556: SetPixelV lHDC, tmpw + 24, tmph + 20, &H7F7D75: SetPixelV lHDC, tmpw + 25, tmph + 20, &HA4A3A1: SetPixelV lHDC, tmpw + 26, tmph + 20, &HC5C5C5: SetPixelV lHDC, tmpw + 27, tmph + 20, &HDADADA: SetPixelV lHDC, tmpw + 28, tmph + 20, &HEDEDED: SetPixelV lHDC, tmpw + 29, tmph + 20, &HFAFAFA:
  SetPixelV lHDC, tmpw + 17, tmph + 21, &HC5C5C5: SetPixelV lHDC, tmpw + 18, tmph + 21, &HC5C5C5: SetPixelV lHDC, tmpw + 19, tmph + 21, &HC6C6C6: SetPixelV lHDC, tmpw + 20, tmph + 21, &HC6C6C6: SetPixelV lHDC, tmpw + 21, tmph + 21, &HC6C6C6: SetPixelV lHDC, tmpw + 22, tmph + 21, &HC9C9C9: SetPixelV lHDC, tmpw + 23, tmph + 21, &HCECECE: SetPixelV lHDC, tmpw + 24, tmph + 21, &HD7D7D7: SetPixelV lHDC, tmpw + 25, tmph + 21, &HE1E1E1: SetPixelV lHDC, tmpw + 26, tmph + 21, &HECECEC: SetPixelV lHDC, tmpw + 27, tmph + 21, &HF6F6F6: SetPixelV lHDC, tmpw + 28, tmph + 21, &HFDFDFD:
  SetPixelV lHDC, tmpw + 17, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 18, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 19, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 20, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 21, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 22, tmph + 22, &HEDEDED: SetPixelV lHDC, tmpw + 23, tmph + 22, &HF0F0F0: SetPixelV lHDC, tmpw + 24, tmph + 22, &HF4F4F4: SetPixelV lHDC, tmpw + 25, tmph + 22, &HFAFAFA: SetPixelV lHDC, tmpw + 26, tmph + 22, &HFDFDFD:
  tmph = 11:     tmph1 = lh - 10:     tmpw = lw - 34
  'Generar lineas intermedias
  APILine 0, tmph, 0, tmph1, &HF7F7F7: APILine 1, tmph, 1, tmph1, &HB0A09E: APILine 2, tmph, 2, tmph1, &H712E13: APILine 3, tmph, 3, tmph1, &HBD5F14:
  APILine 4, tmph, 4, tmph1, &HD17327: APILine 5, tmph, 5, tmph1, &HD47F31: APILine 6, tmph, 6, tmph1, &HD98C3D: APILine 7, tmph, 7, tmph1, &HD9944B:
  APILine 8, tmph, 8, tmph1, &HD7944F: APILine 9, tmph, 9, tmph1, &HDC9C55: APILine 10, tmph, 10, tmph1, &HDC9B57: APILine 11, tmph, 11, tmph1, &HE3A362:
  APILine 12, tmph, 12, tmph1, &HE3A265: APILine 13, tmph, 13, tmph1, &HE2A367: APILine 14, tmph, 14, tmph1, &HE0A165: APILine 15, tmph, 15, tmph1, &HE3A66A:
  APILine 16, tmph, 16, tmph1, &HE3A66A: APILine 17, tmph, 17, tmph1, &HE2A66A:
  APILine tmpw + 17, tmph, tmpw + 17, tmph1, &HE2A66A: APILine tmpw + 18, tmph, tmpw + 18, tmph1, &HE2A66A: APILine tmpw + 19, tmph, tmpw + 19, tmph1, &HE1A464:
  APILine tmpw + 20, tmph, tmpw + 20, tmph1, &HE0A363: APILine tmpw + 21, tmph, tmpw + 21, tmph1, &HE0A363: APILine tmpw + 22, tmph, tmpw + 22, tmph1, &HE1A161
  APILine tmpw + 23, tmph, tmpw + 23, tmph1, &HE09F5B: APILine tmpw + 24, tmph, tmpw + 24, tmph1, &HDE9855: APILine tmpw + 25, tmph, tmpw + 25, tmph1, &HDC9752:
  APILine tmpw + 26, tmph, tmpw + 26, tmph1, &HDB934B: APILine tmpw + 27, tmph, tmpw + 27, tmph1, &HD68D39: APILine tmpw + 28, tmph, tmpw + 28, tmph1, &HD17F2D:
  APILine tmpw + 29, tmph, tmpw + 29, tmph1, &HD67426: APILine tmpw + 30, tmph, tmpw + 30, tmph1, &HC05D13: APILine tmpw + 31, tmph, tmpw + 31, tmph1, &H7C3514:
  APILine tmpw + 32, tmph, tmpw + 32, tmph1, &HAB9B98: APILine tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6:
  'Lineas verticales
  APILine 17, 0, lw - 17, 0, &H450608
  APILine 17, 1, lw - 17, 1, &HF1D4C9
  APILine 17, 2, lw - 17, 2, &HE5C8BD
  APILine 17, 3, lw - 17, 3, &HE8C0A1
  APILine 17, 4, lw - 17, 4, &HE0B898
  APILine 17, 5, lw - 17, 5, &HE3B48E
  APILine 17, 6, lw - 17, 6, &HE0B18B
  APILine 17, 7, lw - 17, 7, &HE9B47F
  APILine 17, 8, lw - 17, 8, &HCE9963
  APILine 17, 9, lw - 17, 9, &HDDA064
  APILine 17, 10, lw - 17, 10, &HE2A66A
  APILine 17, 11, lw - 17, 11, &HE6AC76
  tmph = lh - 22
  APILine 17, tmph + 11, lw - 17, tmph + 11, &HE6AC76
  APILine 17, tmph + 12, lw - 17, tmph + 12, &HF1B681
  APILine 17, tmph + 13, lw - 17, tmph + 13, &HF3BD8A
  APILine 17, tmph + 14, lw - 17, tmph + 14, &HFCC592
  APILine 17, tmph + 15, lw - 17, tmph + 15, &HF8CE97
  APILine 17, tmph + 16, lw - 17, tmph + 16, &HFED59E
  APILine 17, tmph + 17, lw - 17, tmph + 17, &HF7DDA3
  APILine 17, tmph + 18, lw - 17, tmph + 18, &HFFE6AD
  APILine 17, tmph + 19, lw - 17, tmph + 19, &HE9E2C5
  APILine 17, tmph + 20, lw - 17, tmph + 20, &H635D40
  APILine 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
  APILine 17, tmph + 22, lw - 17, tmph + 22, &HECECEC
  Exit Sub

DrawMacOSXButtonHot_Error:
End Sub

Private Sub DrawMacOSXButtonPressed()
  On Error GoTo DrawMacOSXButtonPressed_Error

  Dim lHDC As Long
  lHDC = UserControl.hdc
  'Variable vars (real into code)
  Dim lh As Long, lw As Long
  lh = UserControl.ScaleHeight: lw = UserControl.ScaleWidth
  Dim tmph As Long, tmpw As Long
  Dim tmph1 As Long, tmpw1 As Long
  APIFillRectByCoords hdc, 18, 11, lw - 34, lh - 19, &HCC9B6A
  SetPixelV lHDC, 6, 0, &HFEFEFE: SetPixelV lHDC, 7, 0, &HE5E4E4: SetPixelV lHDC, 8, 0, &HA5A2A2: SetPixelV lHDC, 9, 0, &H675C5C: SetPixelV lHDC, 10, 0, &H422729: SetPixelV lHDC, 11, 0, &H300E0D: SetPixelV lHDC, 12, 0, &H300A09: SetPixelV lHDC, 13, 0, &H2F0908: SetPixelV lHDC, 14, 0, &H330909: SetPixelV lHDC, 15, 0, &H390A0A: SetPixelV lHDC, 16, 0, &H3C0A0A: SetPixelV lHDC, 17, 0, &H3C090A:
  SetPixelV lHDC, 5, 1, &HF0EEEE: SetPixelV lHDC, 6, 1, &H9D888A: SetPixelV lHDC, 7, 1, &H653531: SetPixelV lHDC, 8, 1, &H5A201D: SetPixelV lHDC, 9, 1, &H8D655F: SetPixelV lHDC, 10, 1, &HB99995: SetPixelV lHDC, 11, 1, &HD0B4B2: SetPixelV lHDC, 12, 1, &HD7BEBB: SetPixelV lHDC, 13, 1, &HDDC6C0: SetPixelV lHDC, 14, 1, &HDDC6C0: SetPixelV lHDC, 15, 1, &HDDC7BE: SetPixelV lHDC, 16, 1, &HDDC7BE: SetPixelV lHDC, 17, 1, &HDEC7BE:
  SetPixelV lHDC, 3, 2, &HFEFEFE: SetPixelV lHDC, 4, 2, &HE4E4E4: SetPixelV lHDC, 5, 2, &H6F5C5C: SetPixelV lHDC, 6, 2, &H390A0E: SetPixelV lHDC, 7, 2, &H712E2A: SetPixelV lHDC, 8, 2, &HD6928D: SetPixelV lHDC, 9, 2, &HD8ACA6: SetPixelV lHDC, 10, 2, &HD1B0AC: SetPixelV lHDC, 11, 2, &HD1B5B2: SetPixelV lHDC, 12, 2, &HD0B7B4: SetPixelV lHDC, 13, 2, &HCEB7B1: SetPixelV lHDC, 14, 2, &HCEB7B1: SetPixelV lHDC, 15, 2, &HD2BCB2: SetPixelV lHDC, 16, 2, &HD2BCB2: SetPixelV lHDC, 17, 2, &HD3BCB2:
  SetPixelV lHDC, 3, 3, &HEEEDED: SetPixelV lHDC, 4, 3, &H805858: SetPixelV lHDC, 5, 3, &H6A0D08: SetPixelV lHDC, 6, 3, &H7D1909: SetPixelV lHDC, 7, 3, &HB07B63: SetPixelV lHDC, 8, 3, &HCFA58A: SetPixelV lHDC, 9, 3, &HCDA78E: SetPixelV lHDC, 10, 3, &HD1AB92: SetPixelV lHDC, 11, 3, &HD2AF93: SetPixelV lHDC, 12, 3, &HD3B094: SetPixelV lHDC, 13, 3, &HD0AF93: SetPixelV lHDC, 14, 3, &HD3B296: SetPixelV lHDC, 15, 3, &HD4B49A: SetPixelV lHDC, 16, 3, &HD4B39A: SetPixelV lHDC, 17, 3, &HD4B39A:
  SetPixelV lHDC, 2, 4, &HFBFBFB: SetPixelV lHDC, 3, 4, &H837576: SetPixelV lHDC, 4, 4, &H440C0C: SetPixelV lHDC, 5, 4, &H821D0D: SetPixelV lHDC, 6, 4, &HA94433: SetPixelV lHDC, 7, 4, &HC08B72: SetPixelV lHDC, 8, 4, &HC49A7F: SetPixelV lHDC, 9, 4, &HC6A188: SetPixelV lHDC, 10, 4, &HC7A189: SetPixelV lHDC, 11, 4, &HC6A387: SetPixelV lHDC, 12, 4, &HC8A689: SetPixelV lHDC, 13, 4, &HC9A98C: SetPixelV lHDC, 14, 4, &HC8A88B: SetPixelV lHDC, 15, 4, &HCBAB91: SetPixelV lHDC, 16, 4, &HCCAC92: SetPixelV lHDC, 17, 4, &HCCAC92:
  SetPixelV lHDC, 1, 5, &HFEFEFE: SetPixelV lHDC, 2, 5, &HCAC8C7: SetPixelV lHDC, 3, 5, &H79281D: SetPixelV lHDC, 4, 5, &H7F2409: SetPixelV lHDC, 5, 5, &H8C3809: SetPixelV lHDC, 6, 5, &HBD6D39: SetPixelV lHDC, 7, 5, &HC9986E: SetPixelV lHDC, 8, 5, &HC89D74: SetPixelV lHDC, 9, 5, &HC49A71: SetPixelV lHDC, 10, 5, &HCAA17C: SetPixelV lHDC, 11, 5, &HC6A07A: SetPixelV lHDC, 12, 5, &HCAA480: SetPixelV lHDC, 13, 5, &HCAA582: SetPixelV lHDC, 14, 5, &HCBA584: SetPixelV lHDC, 15, 5, &HCDA989: SetPixelV lHDC, 16, 5, &HCFA98A: SetPixelV lHDC, 17, 5, &HCFA88A:
  SetPixelV lHDC, 1, 6, &HF9F9F9: SetPixelV lHDC, 2, 6, &H756C6B: SetPixelV lHDC, 3, 6, &H76190D: SetPixelV lHDC, 4, 6, &H913416: SetPixelV lHDC, 5, 6, &H9D4916: SetPixelV lHDC, 6, 6, &HBA6A36: SetPixelV lHDC, 7, 6, &HC39268: SetPixelV lHDC, 8, 6, &HC59A71: SetPixelV lHDC, 9, 6, &HC59B72: SetPixelV lHDC, 10, 6, &HC59C77: SetPixelV lHDC, 11, 6, &HC6A07A: SetPixelV lHDC, 12, 6, &HC6A07C: SetPixelV lHDC, 13, 6, &HC7A27F: SetPixelV lHDC, 14, 6, &HCBA584: SetPixelV lHDC, 15, 6, &HCAA686: SetPixelV lHDC, 16, 6, &HCBA586: SetPixelV lHDC, 17, 6, &HCCA587:
  SetPixelV lHDC, 1, 7, &HE8E7E7: SetPixelV lHDC, 2, 7, &H6C3E35: SetPixelV lHDC, 3, 7, &H8A2D09: SetPixelV lHDC, 4, 7, &HA34812: SetPixelV lHDC, 5, 7, &HAB591A: SetPixelV lHDC, 6, 7, &HB46B2B: SetPixelV lHDC, 7, 7, &HC3854A: SetPixelV lHDC, 8, 7, &HD19C64: SetPixelV lHDC, 9, 7, &HCD9C6C: SetPixelV lHDC, 10, 7, &HD1A070: SetPixelV lHDC, 11, 7, &HD2A272: SetPixelV lHDC, 12, 7, &HD2A272: SetPixelV lHDC, 13, 7, &HD6A57A: SetPixelV lHDC, 14, 7, &HD8A77C: SetPixelV lHDC, 15, 7, &HD2A87C: SetPixelV lHDC, 16, 7, &HD2A87C: SetPixelV lHDC, 17, 7, &HD2A77D:
  SetPixelV lHDC, 0, 8, &HFDFDFD: SetPixelV lHDC, 1, 8, &HC7C3C3: SetPixelV lHDC, 2, 8, &H5C2A21: SetPixelV lHDC, 3, 8, &H9C3E15: SetPixelV lHDC, 4, 8, &HB35A22: SetPixelV lHDC, 5, 8, &HB56324: SetPixelV lHDC, 6, 8, &HB66D2D: SetPixelV lHDC, 7, 8, &HB6783D: SetPixelV lHDC, 8, 8, &HB07B44: SetPixelV lHDC, 9, 8, &HB18050: SetPixelV lHDC, 10, 8, &HB58454: SetPixelV lHDC, 11, 8, &HB48554: SetPixelV lHDC, 12, 8, &HB78858: SetPixelV lHDC, 13, 8, &HBA895E: SetPixelV lHDC, 14, 8, &HBB8A5F: SetPixelV lHDC, 15, 8, &HB98E62: SetPixelV lHDC, 16, 8, &HB98E62: SetPixelV lHDC, 17, 8, &HB98E62:
  SetPixelV lHDC, 0, 9, &HFAFAFA: SetPixelV lHDC, 1, 9, &HB4ABA9: SetPixelV lHDC, 2, 9, &H612A14: SetPixelV lHDC, 3, 9, &HA05316: SetPixelV lHDC, 4, 9, &HB36628: SetPixelV lHDC, 5, 9, &HB67132: SetPixelV lHDC, 6, 9, &HB67738: SetPixelV lHDC, 7, 9, &HB98146: SetPixelV lHDC, 8, 9, &HBD864E: SetPixelV lHDC, 9, 9, &HBD894F: SetPixelV lHDC, 10, 9, &HC28D55: SetPixelV lHDC, 11, 9, &HC4905B: SetPixelV lHDC, 12, 9, &HC5905F: SetPixelV lHDC, 13, 9, &HC49161: SetPixelV lHDC, 14, 9, &HC49161: SetPixelV lHDC, 15, 9, &HC69564: SetPixelV lHDC, 16, 9, &HC69564: SetPixelV lHDC, 17, 9, &HC69464:
  SetPixelV lHDC, 0, 10, &HF7F7F7: SetPixelV lHDC, 1, 10, &HA99D9B: SetPixelV lHDC, 2, 10, &H632D17: SetPixelV lHDC, 3, 10, &HA65A1D: SetPixelV lHDC, 4, 10, &HB96C2E: SetPixelV lHDC, 5, 10, &HBC7738: SetPixelV lHDC, 6, 10, &HC18242: SetPixelV lHDC, 7, 10, &HC2894E: SetPixelV lHDC, 8, 10, &HC18A52: SetPixelV lHDC, 9, 10, &HC59157: SetPixelV lHDC, 10, 10, &HC59159: SetPixelV lHDC, 11, 10, &HCC9863: SetPixelV lHDC, 12, 10, &HCC9665: SetPixelV lHDC, 13, 10, &HCB9767: SetPixelV lHDC, 14, 10, &HC99565: SetPixelV lHDC, 15, 10, &HCC9A6A: SetPixelV lHDC, 16, 10, &HCC9A6A: SetPixelV lHDC, 17, 10, &HCC9B6A:
  tmph = lh - 22
  SetPixelV lHDC, 0, tmph + 10, &HF7F7F7: SetPixelV lHDC, 1, tmph + 10, &HA99D9B: SetPixelV lHDC, 2, tmph + 10, &H632D17: SetPixelV lHDC, 3, tmph + 10, &HA65A1D: SetPixelV lHDC, 4, tmph + 10, &HB96C2E: SetPixelV lHDC, 5, tmph + 10, &HBC7738: SetPixelV lHDC, 6, tmph + 10, &HC18242: SetPixelV lHDC, 7, tmph + 10, &HC2894E: SetPixelV lHDC, 8, tmph + 10, &HC18A52: SetPixelV lHDC, 9, tmph + 10, &HC59157: SetPixelV lHDC, 10, tmph + 10, &HC59159: SetPixelV lHDC, 11, tmph + 10, &HCC9863: SetPixelV lHDC, 12, tmph + 10, &HCC9665: SetPixelV lHDC, 13, tmph + 10, &HCB9767: SetPixelV lHDC, 14, tmph + 10, &HC99565: SetPixelV lHDC, 15, tmph + 10, &HCC9A6A: SetPixelV lHDC, 16, tmph + 10, &HCC9A6A: SetPixelV lHDC, 17, tmph + 10, &HCC9B6A:
  SetPixelV lHDC, 0, tmph + 11, &HF5F5F5: SetPixelV lHDC, 1, tmph + 11, &HA59F9A: SetPixelV lHDC, 2, tmph + 11, &H674024: SetPixelV lHDC, 3, tmph + 11, &HAE6827: SetPixelV lHDC, 4, tmph + 11, &HB97231: SetPixelV lHDC, 5, tmph + 11, &HBE8247: SetPixelV lHDC, 6, tmph + 11, &HC0874E: SetPixelV lHDC, 7, tmph + 11, &HC78E56: SetPixelV lHDC, 8, tmph + 11, &HCD9561: SetPixelV lHDC, 9, tmph + 11, &HCB9466: SetPixelV lHDC, 10, tmph + 11, &HCD9A6B: SetPixelV lHDC, 11, tmph + 11, &HC79867: SetPixelV lHDC, 12, tmph + 11, &HCA9B6A: SetPixelV lHDC, 13, tmph + 11, &HCC9D6C: SetPixelV lHDC, 14, tmph + 11, &HCD9D70: SetPixelV lHDC, 15, tmph + 11, &HD0A175: SetPixelV lHDC, 16, tmph + 11, &HD0A175: SetPixelV lHDC, 17, tmph + 11, &HD0A175:
  SetPixelV lHDC, 0, tmph + 12, &HF5F5F5: SetPixelV lHDC, 1, tmph + 12, &HACA7A4: SetPixelV lHDC, 2, tmph + 12, &H755035: SetPixelV lHDC, 3, tmph + 12, &HB77131: SetPixelV lHDC, 4, tmph + 12, &HCB8443: SetPixelV lHDC, 5, tmph + 12, &HC5894E: SetPixelV lHDC, 6, tmph + 12, &HCC935A: SetPixelV lHDC, 7, tmph + 12, &HD29962: SetPixelV lHDC, 8, tmph + 12, &HD69F6A: SetPixelV lHDC, 9, tmph + 12, &HDBA476: SetPixelV lHDC, 10, tmph + 12, &HD6A374: SetPixelV lHDC, 11, tmph + 12, &HD4A574: SetPixelV lHDC, 12, tmph + 12, &HD8A978: SetPixelV lHDC, 13, tmph + 12, &HDAAB7A: SetPixelV lHDC, 14, tmph + 12, &HDAAA7D: SetPixelV lHDC, 15, tmph + 12, &HDBAB7F: SetPixelV lHDC, 16, tmph + 12, &HDAAA7F: SetPixelV lHDC, 17, tmph + 12, &HDAAA7F:
  SetPixelV lHDC, 0, tmph + 13, &HF7F7F7: SetPixelV lHDC, 1, tmph + 13, &HC0C0BF: SetPixelV lHDC, 2, tmph + 13, &H63574B: SetPixelV lHDC, 3, tmph + 13, &HAC7036: SetPixelV lHDC, 4, tmph + 13, &HC2854A: SetPixelV lHDC, 5, tmph + 13, &HCF955E: SetPixelV lHDC, 6, tmph + 13, &HD29B66: SetPixelV lHDC, 7, tmph + 13, &HD1A26E: SetPixelV lHDC, 8, tmph + 13, &HD8A776: SetPixelV lHDC, 9, tmph + 13, &HDBA878: SetPixelV lHDC, 10, tmph + 13, &HDFAC7C: SetPixelV lHDC, 11, tmph + 13, &HDBAF7D: SetPixelV lHDC, 12, tmph + 13, &HDDAF81: SetPixelV lHDC, 13, tmph + 13, &HDEB183: SetPixelV lHDC, 14, tmph + 13, &HDDAF84: SetPixelV lHDC, 15, tmph + 13, &HDEB087: SetPixelV lHDC, 16, tmph + 13, &HDEB087: SetPixelV lHDC, 17, tmph + 13, &HDCB087:
  SetPixelV lHDC, 0, tmph + 14, &HFBFBFB: SetPixelV lHDC, 1, tmph + 14, &HE1E1E1: SetPixelV lHDC, 2, tmph + 14, &H7C7269: SetPixelV lHDC, 3, tmph + 14, &HA26830: SetPixelV lHDC, 4, tmph + 14, &HC6884E: SetPixelV lHDC, 5, tmph + 14, &HD0965F: SetPixelV lHDC, 6, tmph + 14, &HDAA26E: SetPixelV lHDC, 7, tmph + 14, &HD9AA75: SetPixelV lHDC, 8, tmph + 14, &HDBAA79: SetPixelV lHDC, 9, tmph + 14, &HE2AF7F: SetPixelV lHDC, 10, tmph + 14, &HE6B484: SetPixelV lHDC, 11, tmph + 14, &HE2B684: SetPixelV lHDC, 12, tmph + 14, &HE3B588: SetPixelV lHDC, 13, tmph + 14, &HE2B587: SetPixelV lHDC, 14, tmph + 14, &HE2B48A: SetPixelV lHDC, 15, tmph + 14, &HE5B78E: SetPixelV lHDC, 16, tmph + 14, &HE5B78E: SetPixelV lHDC, 17, tmph + 14, &HE4B88E:
  SetPixelV lHDC, 0, tmph + 15, &HFEFEFE: SetPixelV lHDC, 1, tmph + 15, &HEDEDED: SetPixelV lHDC, 2, tmph + 15, &H9E9C9C: SetPixelV lHDC, 3, tmph + 15, &H766051: SetPixelV lHDC, 4, tmph + 15, &HAD8666: SetPixelV lHDC, 5, tmph + 15, &HD49A61: SetPixelV lHDC, 6, tmph + 15, &HE0A66D: SetPixelV lHDC, 7, tmph + 15, &HE3B17C: SetPixelV lHDC, 8, tmph + 15, &HE0B380: SetPixelV lHDC, 9, tmph + 15, &HE0B587: SetPixelV lHDC, 10, tmph + 15, &HE2BC8C: SetPixelV lHDC, 11, tmph + 15, &HE0BB8B: SetPixelV lHDC, 12, tmph + 15, &HE0BC8B: SetPixelV lHDC, 13, tmph + 15, &HE3BD92: SetPixelV lHDC, 14, tmph + 15, &HE2BC91: SetPixelV lHDC, 15, tmph + 15, &HE2BF93: SetPixelV lHDC, 16, tmph + 15, &HE1BE93: SetPixelV lHDC, 17, tmph + 15, &HE1BF93:
  SetPixelV lHDC, 1, tmph + 16, &HF6F6F6: SetPixelV lHDC, 2, tmph + 16, &HD5D5D5: SetPixelV lHDC, 3, tmph + 16, &H86766C: SetPixelV lHDC, 4, tmph + 16, &H856144: SetPixelV lHDC, 5, tmph + 16, &HD59C63: SetPixelV lHDC, 6, tmph + 16, &HE5AB71: SetPixelV lHDC, 7, tmph + 16, &HE5B37E: SetPixelV lHDC, 8, tmph + 16, &HE7BB88: SetPixelV lHDC, 9, tmph + 16, &HE7BF91: SetPixelV lHDC, 10, tmph + 16, &HE3BC8D: SetPixelV lHDC, 11, tmph + 16, &HE7C392: SetPixelV lHDC, 12, tmph + 16, &HE7C392: SetPixelV lHDC, 13, tmph + 16, &HE8C398: SetPixelV lHDC, 14, tmph + 16, &HE8C499: SetPixelV lHDC, 15, tmph + 16, &HE8C599: SetPixelV lHDC, 16, tmph + 16, &HE8C599: SetPixelV lHDC, 17, tmph + 16, &HE7C699:
  SetPixelV lHDC, 1, tmph + 17, &HFDFDFD: SetPixelV lHDC, 2, tmph + 17, &HEDEDED: SetPixelV lHDC, 3, tmph + 17, &HBDBDBD: SetPixelV lHDC, 4, tmph + 17, &H676767: SetPixelV lHDC, 5, tmph + 17, &H71604C: SetPixelV lHDC, 6, tmph + 17, &HBEA17D: SetPixelV lHDC, 7, tmph + 17, &HDAB381: SetPixelV lHDC, 8, tmph + 17, &HE5BE8C: SetPixelV lHDC, 9, tmph + 17, &HE1C18F: SetPixelV lHDC, 10, tmph + 17, &HE4C895: SetPixelV lHDC, 11, tmph + 17, &HDFCA98: SetPixelV lHDC, 12, tmph + 17, &HE2CE9B: SetPixelV lHDC, 13, tmph + 17, &HE2CE9B: SetPixelV lHDC, 14, tmph + 17, &HE2CE9B: SetPixelV lHDC, 15, tmph + 17, &HE2CD9D: SetPixelV lHDC, 16, tmph + 17, &HE2CC9D: SetPixelV lHDC, 17, tmph + 17, &HE2CC9D:
  SetPixelV lHDC, 2, tmph + 18, &HF9F9F9: SetPixelV lHDC, 3, tmph + 18, &HE6E6E6: SetPixelV lHDC, 4, tmph + 18, &HB9B9B9: SetPixelV lHDC, 5, tmph + 18, &H7A7163: SetPixelV lHDC, 6, tmph + 18, &H776043: SetPixelV lHDC, 7, tmph + 18, &HAB885B: SetPixelV lHDC, 8, tmph + 18, &HDDB888: SetPixelV lHDC, 9, tmph + 18, &HE6C796: SetPixelV lHDC, 10, tmph + 18, &HE8CD9A: SetPixelV lHDC, 11, tmph + 18, &HE5D19E: SetPixelV lHDC, 12, tmph + 18, &HE9D6A3: SetPixelV lHDC, 13, tmph + 18, &HE9D6A5: SetPixelV lHDC, 14, tmph + 18, &HE9D6A3: SetPixelV lHDC, 15, tmph + 18, &HE9D5A6: SetPixelV lHDC, 16, tmph + 18, &HE9D5A6: SetPixelV lHDC, 17, tmph + 18, &HE9D5A6:
  SetPixelV lHDC, 2, tmph + 19, &HFEFEFE: SetPixelV lHDC, 3, tmph + 19, &HF8F8F8: SetPixelV lHDC, 4, tmph + 19, &HE6E6E6: SetPixelV lHDC, 5, tmph + 19, &HC8C8C8: SetPixelV lHDC, 6, tmph + 19, &H8C8C8C: SetPixelV lHDC, 7, tmph + 19, &H61605E: SetPixelV lHDC, 8, tmph + 19, &H656059: SetPixelV lHDC, 9, tmph + 19, &H857C6D: SetPixelV lHDC, 10, tmph + 19, &HA59C87: SetPixelV lHDC, 11, tmph + 19, &HC8C1A8: SetPixelV lHDC, 12, tmph + 19, &HD1CAB0: SetPixelV lHDC, 13, tmph + 19, &HD5CFB5: SetPixelV lHDC, 14, tmph + 19, &HD6D1B6: SetPixelV lHDC, 15, tmph + 19, &HD7D2BA: SetPixelV lHDC, 16, tmph + 19, &HD7D1BA: SetPixelV lHDC, 17, tmph + 19, &HD7D2BA:
  SetPixelV lHDC, 3, tmph + 20, &HFEFEFE: SetPixelV lHDC, 4, tmph + 20, &HF9F9F9: SetPixelV lHDC, 5, tmph + 20, &HECECEC: SetPixelV lHDC, 6, tmph + 20, &HDADADA: SetPixelV lHDC, 7, tmph + 20, &HC1C1C1: SetPixelV lHDC, 8, tmph + 20, &H9C9B99: SetPixelV lHDC, 9, tmph + 20, &H7D7A73: SetPixelV lHDC, 10, tmph + 20, &H635E50: SetPixelV lHDC, 11, tmph + 20, &H58533F: SetPixelV lHDC, 12, tmph + 20, &H554F39: SetPixelV lHDC, 13, tmph + 20, &H514D36: SetPixelV lHDC, 14, tmph + 20, &H554F37: SetPixelV lHDC, 15, tmph + 20, &H57523A: SetPixelV lHDC, 16, tmph + 20, &H5A563D: SetPixelV lHDC, 17, tmph + 20, &H5A563E:
  SetPixelV lHDC, 5, tmph + 21, &HFCFCFC: SetPixelV lHDC, 6, tmph + 21, &HF5F5F5: SetPixelV lHDC, 7, tmph + 21, &HEBEBEB: SetPixelV lHDC, 8, tmph + 21, &HE1E1E1: SetPixelV lHDC, 9, tmph + 21, &HD6D6D6: SetPixelV lHDC, 10, tmph + 21, &HCECECE: SetPixelV lHDC, 11, tmph + 21, &HC9C9C9: SetPixelV lHDC, 12, tmph + 21, &HC7C7C7: SetPixelV lHDC, 13, tmph + 21, &HC7C7C7: SetPixelV lHDC, 14, tmph + 21, &HC6C6C6: SetPixelV lHDC, 15, tmph + 21, &HC6C6C6: SetPixelV lHDC, 16, tmph + 21, &HC5C5C5: SetPixelV lHDC, 17, tmph + 21, &HC5C5C5:
  SetPixelV lHDC, 7, tmph + 22, &HFDFDFD: SetPixelV lHDC, 8, tmph + 22, &HF9F9F9: SetPixelV lHDC, 9, tmph + 22, &HF4F4F4: SetPixelV lHDC, 10, tmph + 22, &HF0F0F0: SetPixelV lHDC, 11, tmph + 22, &HEEEEEE: SetPixelV lHDC, 12, tmph + 22, &HEDEDED: SetPixelV lHDC, 13, tmph + 22, &HECECEC: SetPixelV lHDC, 14, tmph + 22, &HECECEC: SetPixelV lHDC, 15, tmph + 22, &HECECEC: SetPixelV lHDC, 16, tmph + 22, &HECECEC: SetPixelV lHDC, 17, tmph + 22, &HECECEC:
  tmpw = lw - 34
  SetPixelV lHDC, tmpw + 17, 0, &H3C090A: SetPixelV lHDC, tmpw + 18, 0, &H3C090A: SetPixelV lHDC, tmpw + 19, 0, &H340A0A: SetPixelV lHDC, tmpw + 20, 0, &H300A09: SetPixelV lHDC, tmpw + 21, 0, &H2F080A: SetPixelV lHDC, tmpw + 22, 0, &H341011: SetPixelV lHDC, tmpw + 23, 0, &H3E2526: SetPixelV lHDC, tmpw + 24, 0, &H5A4C4C: SetPixelV lHDC, tmpw + 25, 0, &H9E9B9B: SetPixelV lHDC, tmpw + 26, 0, &HEEEEEE: SetPixelV lHDC, tmpw + 34, 0, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 1, &HDEC7BE: SetPixelV lHDC, tmpw + 18, 1, &HDEC7BE: SetPixelV lHDC, tmpw + 19, 1, &HDBC6C1: SetPixelV lHDC, tmpw + 20, 1, &HD9C4BF: SetPixelV lHDC, tmpw + 21, 1, &HD7C1B9: SetPixelV lHDC, tmpw + 22, 1, &HD3B5AF: SetPixelV lHDC, tmpw + 23, 1, &HBE9F97: SetPixelV lHDC, tmpw + 24, 1, &H9B6A65: SetPixelV lHDC, tmpw + 25, 1, &H65231E: SetPixelV lHDC, tmpw + 26, 1, &H642A26: SetPixelV lHDC, tmpw + 27, 1, &HA59696: SetPixelV lHDC, tmpw + 28, 1, &HF7F7F7: SetPixelV lHDC, tmpw + 34, 1, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 2, &HD3BCB2: SetPixelV lHDC, tmpw + 18, 2, &HD3BCB2: SetPixelV lHDC, tmpw + 19, 2, &HCDB8B3: SetPixelV lHDC, tmpw + 20, 2, &HCBB6B1: SetPixelV lHDC, tmpw + 21, 2, &HD0BBB2: SetPixelV lHDC, tmpw + 22, 2, &HD0B2AC: SetPixelV lHDC, tmpw + 23, 2, &HD6B6AF: SetPixelV lHDC, tmpw + 24, 2, &HDCABA6: SetPixelV lHDC, tmpw + 25, 2, &HDC9691: SetPixelV lHDC, tmpw + 26, 2, &H732E29: SetPixelV lHDC, tmpw + 27, 2, &H380A0A: SetPixelV lHDC, tmpw + 28, 2, &H6A5556: SetPixelV lHDC, tmpw + 29, 2, &HEAEBEA: SetPixelV lHDC, tmpw + 34, 2, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 3, &HD4B39A: SetPixelV lHDC, tmpw + 18, 3, &HD4B39A: SetPixelV lHDC, tmpw + 19, 3, &HD1B294: SetPixelV lHDC, tmpw + 20, 3, &HD0B193: SetPixelV lHDC, tmpw + 21, 3, &HD0AE91: SetPixelV lHDC, tmpw + 22, 3, &HD4B296: SetPixelV lHDC, tmpw + 23, 3, &HCBAA8F: SetPixelV lHDC, tmpw + 24, 3, &HCBAA8F: SetPixelV lHDC, tmpw + 25, 3, &HCCA38B: SetPixelV lHDC, tmpw + 26, 3, &HB77E68: SetPixelV lHDC, tmpw + 27, 3, &H811B09: SetPixelV lHDC, tmpw + 28, 3, &H720E08: SetPixelV lHDC, tmpw + 29, 3, &H7D5051: SetPixelV lHDC, tmpw + 30, 3, &HEFEEEE: SetPixelV lHDC, tmpw + 34, 3, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 4, &HCCAC92: SetPixelV lHDC, tmpw + 18, 4, &HCCAC91: SetPixelV lHDC, tmpw + 19, 4, &HC6A889: SetPixelV lHDC, tmpw + 20, 4, &HC7A98A: SetPixelV lHDC, tmpw + 21, 4, &HC7A589: SetPixelV lHDC, tmpw + 22, 4, &HC4A185: SetPixelV lHDC, tmpw + 23, 4, &HC6A58A: SetPixelV lHDC, tmpw + 24, 4, &HBF9E83: SetPixelV lHDC, tmpw + 25, 4, &HC39A82: SetPixelV lHDC, tmpw + 26, 4, &HC58C76: SetPixelV lHDC, tmpw + 27, 4, &HA9432F: SetPixelV lHDC, tmpw + 28, 4, &H861F0C: SetPixelV lHDC, tmpw + 29, 4, &H460B0C: SetPixelV lHDC, tmpw + 30, 4, &H7B6B6C: SetPixelV lHDC, tmpw + 31, 4, &HFAFAFA: SetPixelV lHDC, tmpw + 34, 4, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 5, &HCFA88A: SetPixelV lHDC, tmpw + 18, 5, &HCFA889: SetPixelV lHDC, tmpw + 19, 5, &HCBA683: SetPixelV lHDC, tmpw + 20, 5, &HC9A481: SetPixelV lHDC, tmpw + 21, 5, &HCCA480: SetPixelV lHDC, tmpw + 22, 5, &HCEA280: SetPixelV lHDC, tmpw + 23, 5, &HCCA379: SetPixelV lHDC, tmpw + 24, 5, &HCA9E74: SetPixelV lHDC, tmpw + 25, 5, &HC69971: SetPixelV lHDC, tmpw + 26, 5, &HC89870: SetPixelV lHDC, tmpw + 27, 5, &HB46A34: SetPixelV lHDC, tmpw + 28, 5, &H90380A: SetPixelV lHDC, tmpw + 29, 5, &H892509: SetPixelV lHDC, tmpw + 30, 5, &H8A251B: SetPixelV lHDC, tmpw + 31, 5, &HC4C2C2: SetPixelV lHDC, tmpw + 32, 5, &HFEFEFE: SetPixelV lHDC, tmpw + 34, 5, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 6, &HCCA587: SetPixelV lHDC, tmpw + 18, 6, &HCCA586: SetPixelV lHDC, tmpw + 19, 6, &HC9A481: SetPixelV lHDC, tmpw + 20, 6, &HC9A481: SetPixelV lHDC, tmpw + 21, 6, &HC7A07C: SetPixelV lHDC, tmpw + 22, 6, &HCCA17E: SetPixelV lHDC, tmpw + 23, 6, &HC79F74: SetPixelV lHDC, tmpw + 24, 6, &HC69A70: SetPixelV lHDC, tmpw + 25, 6, &HC59870: SetPixelV lHDC, tmpw + 26, 6, &HC2926A: SetPixelV lHDC, tmpw + 27, 6, &HB96F39: SetPixelV lHDC, tmpw + 28, 6, &HA04814: SetPixelV lHDC, tmpw + 29, 6, &H973215: SetPixelV lHDC, tmpw + 30, 6, &H831A0F: SetPixelV lHDC, tmpw + 31, 6, &H6E6966: SetPixelV lHDC, tmpw + 32, 6, &HF8F8F8: SetPixelV lHDC, tmpw + 34, 6, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 7, &HD2A77D: SetPixelV lHDC, tmpw + 18, 7, &HD3A77C: SetPixelV lHDC, tmpw + 19, 7, &HD8AA7D: SetPixelV lHDC, tmpw + 20, 7, &HD2A376: SetPixelV lHDC, tmpw + 21, 7, &HD1A373: SetPixelV lHDC, tmpw + 22, 7, &HCEA070: SetPixelV lHDC, tmpw + 23, 7, &HD2A06F: SetPixelV lHDC, tmpw + 24, 7, &HD19D68: SetPixelV lHDC, tmpw + 25, 7, &HD09A65: SetPixelV lHDC, tmpw + 26, 7, &HC2864F: SetPixelV lHDC, tmpw + 27, 7, &HAE6927: SetPixelV lHDC, tmpw + 28, 7, &HA95A19: SetPixelV lHDC, tmpw + 29, 7, &HA44A10: SetPixelV lHDC, tmpw + 30, 7, &H8B2E09: SetPixelV lHDC, tmpw + 31, 7, &H6B3E34: SetPixelV lHDC, tmpw + 32, 7, &HE7E6E6: SetPixelV lHDC, tmpw + 34, 7, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 8, &HB98E62: SetPixelV lHDC, tmpw + 18, 8, &HBA8E62: SetPixelV lHDC, tmpw + 19, 8, &HB98B5E: SetPixelV lHDC, tmpw + 20, 8, &HB98B5E: SetPixelV lHDC, tmpw + 21, 8, &HB68858: SetPixelV lHDC, tmpw + 22, 8, &HB48656: SetPixelV lHDC, tmpw + 23, 8, &HB58452: SetPixelV lHDC, tmpw + 24, 8, &HB5814C: SetPixelV lHDC, tmpw + 25, 8, &HB07A46: SetPixelV lHDC, tmpw + 26, 8, &HB2773F: SetPixelV lHDC, tmpw + 27, 8, &HB36E2C: SetPixelV lHDC, tmpw + 28, 8, &HB26221: SetPixelV lHDC, tmpw + 29, 8, &HB35A20: SetPixelV lHDC, tmpw + 30, 8, &H9C3E11: SetPixelV lHDC, tmpw + 31, 8, &H5C2A1F: SetPixelV lHDC, tmpw + 32, 8, &HC4C1C0: SetPixelV lHDC, tmpw + 33, 8, &HFDFDFD: SetPixelV lHDC, tmpw + 34, 8, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 9, &HC69464: SetPixelV lHDC, tmpw + 18, 9, &HC69564: SetPixelV lHDC, tmpw + 19, 9, &HC3925E: SetPixelV lHDC, tmpw + 20, 9, &HC3915D: SetPixelV lHDC, tmpw + 21, 9, &HC3925E: SetPixelV lHDC, tmpw + 22, 9, &HC38F5B: SetPixelV lHDC, tmpw + 23, 9, &HC28D55: SetPixelV lHDC, tmpw + 24, 9, &HC08751: SetPixelV lHDC, tmpw + 25, 9, &HBC844C: SetPixelV lHDC, tmpw + 26, 9, &HBC8147: SetPixelV lHDC, tmpw + 27, 9, &HB57936: SetPixelV lHDC, tmpw + 28, 9, &HB3702D: SetPixelV lHDC, tmpw + 29, 9, &HB56626: SetPixelV lHDC, tmpw + 30, 9, &HA25115: SetPixelV lHDC, tmpw + 31, 9, &H662D12: SetPixelV lHDC, tmpw + 32, 9, &HAEA3A1: SetPixelV lHDC, tmpw + 33, 9, &HF9F9F9: SetPixelV lHDC, tmpw + 34, 9, &HFFFFFFFF:
  SetPixelV lHDC, tmpw + 17, 10, &HCC9B6A: SetPixelV lHDC, tmpw + 18, 10, &HCC9B6A: SetPixelV lHDC, tmpw + 19, 10, &HCA9864: SetPixelV lHDC, tmpw + 20, 10, &HC99763: SetPixelV lHDC, tmpw + 21, 10, &HC99763: SetPixelV lHDC, tmpw + 22, 10, &HCA9562: SetPixelV lHDC, tmpw + 23, 10, &HC9945D: SetPixelV lHDC, tmpw + 24, 10, &HC68E57: SetPixelV lHDC, tmpw + 25, 10, &HC48C55: SetPixelV lHDC, tmpw + 26, 10, &HC3884E: SetPixelV lHDC, tmpw + 27, 10, &HBE823E: SetPixelV lHDC, tmpw + 28, 10, &HB97634: SetPixelV lHDC, tmpw + 29, 10, &HBD6D2D: SetPixelV lHDC, tmpw + 30, 10, &HA8581C: SetPixelV lHDC, tmpw + 31, 10, &H6D3319: SetPixelV lHDC, tmpw + 32, 10, &HA49794: SetPixelV lHDC, tmpw + 33, 10, &HF6F6F6: SetPixelV lHDC, tmpw + 34, 10, &HFFFFFFFF:
  tmph = lh - 22
  tmpw = lw - 34
  SetPixelV lHDC, tmpw + 17, tmph + 10, &HCC9B6A: SetPixelV lHDC, tmpw + 18, tmph + 10, &HCC9B6A: SetPixelV lHDC, tmpw + 19, tmph + 10, &HCA9864: SetPixelV lHDC, tmpw + 20, tmph + 10, &HC99763: SetPixelV lHDC, tmpw + 21, tmph + 10, &HC99763: SetPixelV lHDC, tmpw + 22, tmph + 10, &HCA9562: SetPixelV lHDC, tmpw + 23, tmph + 10, &HC9945D: SetPixelV lHDC, tmpw + 24, tmph + 10, &HC68E57: SetPixelV lHDC, tmpw + 25, tmph + 10, &HC48C55: SetPixelV lHDC, tmpw + 26, tmph + 10, &HC3884E: SetPixelV lHDC, tmpw + 27, tmph + 10, &HBE823E: SetPixelV lHDC, tmpw + 28, tmph + 10, &HB97634: SetPixelV lHDC, tmpw + 29, tmph + 10, &HBD6D2D: SetPixelV lHDC, tmpw + 30, tmph + 10, &HA8581C: SetPixelV lHDC, tmpw + 31, tmph + 10, &H6D3319: SetPixelV lHDC, tmpw + 32, tmph + 10, &HA49794: SetPixelV lHDC, tmpw + 33, tmph + 10, &HF6F6F6:
  SetPixelV lHDC, tmpw + 17, tmph + 11, &HD0A175: SetPixelV lHDC, tmpw + 18, tmph + 11, &HD0A175: SetPixelV lHDC, tmpw + 19, tmph + 11, &HCC9D6D: SetPixelV lHDC, tmpw + 20, tmph + 11, &HCE9B6C: SetPixelV lHDC, tmpw + 21, tmph + 11, &HCB9A6A: SetPixelV lHDC, tmpw + 22, tmph + 11, &HCD996A: SetPixelV lHDC, tmpw + 23, tmph + 11, &HCA9666: SetPixelV lHDC, tmpw + 24, tmph + 11, &HCF9865: SetPixelV lHDC, tmpw + 25, tmph + 11, &HCA9460: SetPixelV lHDC, tmpw + 26, tmph + 11, &HC78F57: SetPixelV lHDC, tmpw + 27, tmph + 11, &HC1864B: SetPixelV lHDC, tmpw + 28, tmph + 11, &HC08143: SetPixelV lHDC, tmpw + 29, tmph + 11, &HB7712E: SetPixelV lHDC, tmpw + 30, tmph + 11, &HB16A28: SetPixelV lHDC, tmpw + 31, tmph + 11, &H694321: SetPixelV lHDC, tmpw + 32, tmph + 11, &HA59F9B: SetPixelV lHDC, tmpw + 33, tmph + 11, &HF4F4F4:
  SetPixelV lHDC, tmpw + 17, tmph + 12, &HDAAA7F: SetPixelV lHDC, tmpw + 18, tmph + 12, &HD9AB7E: SetPixelV lHDC, tmpw + 19, tmph + 12, &HDBAC7C: SetPixelV lHDC, tmpw + 20, tmph + 12, &HDDAA7B: SetPixelV lHDC, tmpw + 21, tmph + 12, &HDAA979: SetPixelV lHDC, tmpw + 22, tmph + 12, &HDAA677: SetPixelV lHDC, tmpw + 23, tmph + 12, &HD8A474: SetPixelV lHDC, tmpw + 24, tmph + 12, &HDBA471: SetPixelV lHDC, tmpw + 25, tmph + 12, &HD49F6A: SetPixelV lHDC, tmpw + 26, tmph + 12, &HD09861: SetPixelV lHDC, tmpw + 27, tmph + 12, &HD0955A: SetPixelV lHDC, tmpw + 28, tmph + 12, &HCC8D4F: SetPixelV lHDC, tmpw + 29, tmph + 12, &HCA8441: SetPixelV lHDC, tmpw + 30, tmph + 12, &HBB7532: SetPixelV lHDC, tmpw + 31, tmph + 12, &H7B5434: SetPixelV lHDC, tmpw + 32, tmph + 12, &HB1ACAA: SetPixelV lHDC, tmpw + 33, tmph + 12, &HF5F5F5:
  SetPixelV lHDC, tmpw + 17, tmph + 13, &HDCB087: SetPixelV lHDC, tmpw + 18, tmph + 13, &HDCB087: SetPixelV lHDC, tmpw + 19, tmph + 13, &HDBAF81: SetPixelV lHDC, tmpw + 20, tmph + 13, &HDEAF82: SetPixelV lHDC, tmpw + 21, tmph + 13, &HDCAF81: SetPixelV lHDC, tmpw + 22, tmph + 13, &HDDAD7F: SetPixelV lHDC, tmpw + 23, tmph + 13, &HDBAC7B: SetPixelV lHDC, tmpw + 24, tmph + 13, &HDDAA7A: SetPixelV lHDC, tmpw + 25, tmph + 13, &HD9A775: SetPixelV lHDC, tmpw + 26, tmph + 13, &HD7A26E: SetPixelV lHDC, tmpw + 27, tmph + 13, &HCE9961: SetPixelV lHDC, tmpw + 28, tmph + 13, &HCC945C: SetPixelV lHDC, tmpw + 29, tmph + 13, &HC2854D: SetPixelV lHDC, tmpw + 30, tmph + 13, &HAF7239: SetPixelV lHDC, tmpw + 31, tmph + 13, &H695C4F: SetPixelV lHDC, tmpw + 32, tmph + 13, &HD5D5D5: SetPixelV lHDC, tmpw + 33, tmph + 13, &HF8F8F8:
  SetPixelV lHDC, tmpw + 17, tmph + 14, &HE4B88E: SetPixelV lHDC, tmpw + 18, tmph + 14, &HE3B88E: SetPixelV lHDC, tmpw + 19, tmph + 14, &HE0B486: SetPixelV lHDC, tmpw + 20, tmph + 14, &HE4B689: SetPixelV lHDC, tmpw + 21, tmph + 14, &HE2B587: SetPixelV lHDC, tmpw + 22, tmph + 14, &HE5B588: SetPixelV lHDC, tmpw + 23, tmph + 14, &HE3B483: SetPixelV lHDC, tmpw + 24, tmph + 14, &HE2AF7F: SetPixelV lHDC, tmpw + 25, tmph + 14, &HDEAC7A: SetPixelV lHDC, tmpw + 26, tmph + 14, &HDEAA75: SetPixelV lHDC, tmpw + 27, tmph + 14, &HD8A36B: SetPixelV lHDC, tmpw + 28, tmph + 14, &HD09860: SetPixelV lHDC, tmpw + 29, tmph + 14, &HC58850: SetPixelV lHDC, tmpw + 30, tmph + 14, &HA56930: SetPixelV lHDC, tmpw + 31, tmph + 14, &H7B746C: SetPixelV lHDC, tmpw + 32, tmph + 14, &HE8E8E8: SetPixelV lHDC, tmpw + 33, tmph + 14, &HFDFDFD:
  SetPixelV lHDC, tmpw + 17, tmph + 15, &HE1BF93: SetPixelV lHDC, tmpw + 18, tmph + 15, &HE2BE93: SetPixelV lHDC, tmpw + 19, tmph + 15, &HE2BF91: SetPixelV lHDC, tmpw + 20, tmph + 15, &HE1BD8F: SetPixelV lHDC, tmpw + 21, tmph + 15, &HE0BC8E: SetPixelV lHDC, tmpw + 22, tmph + 15, &HE2BC8E: SetPixelV lHDC, tmpw + 23, tmph + 15, &HE4BD8C: SetPixelV lHDC, tmpw + 24, tmph + 15, &HE0B685: SetPixelV lHDC, tmpw + 25, tmph + 15, &HDCB07E: SetPixelV lHDC, tmpw + 26, tmph + 15, &HE1AF7C: SetPixelV lHDC, tmpw + 27, tmph + 15, &HDEA66E: SetPixelV lHDC, tmpw + 28, tmph + 15, &HD19962: SetPixelV lHDC, tmpw + 29, tmph + 15, &HAD875D: SetPixelV lHDC, tmpw + 30, tmph + 15, &H7D6851: SetPixelV lHDC, tmpw + 31, tmph + 15, &HB9B9B9: SetPixelV lHDC, tmpw + 32, tmph + 15, &HF1F1F1:
  SetPixelV lHDC, tmpw + 17, tmph + 16, &HE7C699: SetPixelV lHDC, tmpw + 18, tmph + 16, &HE8C69A: SetPixelV lHDC, tmpw + 19, tmph + 16, &HE7C496: SetPixelV lHDC, tmpw + 20, tmph + 16, &HE8C597: SetPixelV lHDC, tmpw + 21, tmph + 16, &HE5C294: SetPixelV lHDC, tmpw + 22, tmph + 16, &HE8C194: SetPixelV lHDC, tmpw + 23, tmph + 16, &HE6BF8E: SetPixelV lHDC, tmpw + 24, tmph + 16, &HE7BC8C: SetPixelV lHDC, tmpw + 25, tmph + 16, &HE7BB8A: SetPixelV lHDC, tmpw + 26, tmph + 16, &HE5B37F: SetPixelV lHDC, tmpw + 27, tmph + 16, &HE1A971: SetPixelV lHDC, tmpw + 28, tmph + 16, &HD79F67: SetPixelV lHDC, tmpw + 29, tmph + 16, &H8E6A40: SetPixelV lHDC, tmpw + 30, tmph + 16, &H8A8076: SetPixelV lHDC, tmpw + 31, tmph + 16, &HE2E2E2: SetPixelV lHDC, tmpw + 32, tmph + 16, &HF9F9F9:
  SetPixelV lHDC, tmpw + 17, tmph + 17, &HE2CC9D: SetPixelV lHDC, tmpw + 18, tmph + 17, &HE2CC9E: SetPixelV lHDC, tmpw + 19, tmph + 17, &HE3CF9C: SetPixelV lHDC, tmpw + 20, tmph + 17, &HDFCA98: SetPixelV lHDC, tmpw + 21, tmph + 17, &HE2CD9C: SetPixelV lHDC, tmpw + 22, tmph + 17, &HE4CC9C: SetPixelV lHDC, tmpw + 23, tmph + 17, &HE1C491: SetPixelV lHDC, tmpw + 24, tmph + 17, &HE1C18F: SetPixelV lHDC, tmpw + 25, tmph + 17, &HE4BD8C: SetPixelV lHDC, tmpw + 26, tmph + 17, &HDAB285: SetPixelV lHDC, tmpw + 27, tmph + 17, &HC7A582: SetPixelV lHDC, tmpw + 28, tmph + 17, &H806A56: SetPixelV lHDC, tmpw + 29, tmph + 17, &H676565: SetPixelV lHDC, tmpw + 30, tmph + 17, &HD2D2D2: SetPixelV lHDC, tmpw + 31, tmph + 17, &HF2F2F2: SetPixelV lHDC, tmpw + 32, tmph + 17, &HFEFEFE:
  SetPixelV lHDC, tmpw + 17, tmph + 18, &HE9D5A6: SetPixelV lHDC, tmpw + 18, tmph + 18, &HE9D5A6: SetPixelV lHDC, tmpw + 19, tmph + 18, &HE9D6A3: SetPixelV lHDC, tmpw + 20, tmph + 18, &HE9D7A6: SetPixelV lHDC, tmpw + 21, tmph + 18, &HE2CC9B: SetPixelV lHDC, tmpw + 22, tmph + 18, &HE8D0A0: SetPixelV lHDC, tmpw + 23, tmph + 18, &HE8CF9C: SetPixelV lHDC, tmpw + 24, tmph + 18, &HE7C795: SetPixelV lHDC, tmpw + 25, tmph + 18, &HE2BC8B: SetPixelV lHDC, tmpw + 26, tmph + 18, &HB68F63: SetPixelV lHDC, tmpw + 27, tmph + 18, &H886948: SetPixelV lHDC, tmpw + 28, tmph + 18, &H786A5D: SetPixelV lHDC, tmpw + 29, tmph + 18, &HC7C7C7: SetPixelV lHDC, tmpw + 30, tmph + 18, &HEBEBEB: SetPixelV lHDC, tmpw + 31, tmph + 18, &HFCFCFC:
  SetPixelV lHDC, tmpw + 17, tmph + 19, &HD7D2BA: SetPixelV lHDC, tmpw + 18, tmph + 19, &HD7D2BA: SetPixelV lHDC, tmpw + 19, tmph + 19, &HD7D1B9: SetPixelV lHDC, tmpw + 20, tmph + 19, &HD5CEB6: SetPixelV lHDC, tmpw + 21, tmph + 19, &HDBD3BB: SetPixelV lHDC, tmpw + 22, tmph + 19, &HC9C1AA: SetPixelV lHDC, tmpw + 23, tmph + 19, &HA9A28B: SetPixelV lHDC, tmpw + 24, tmph + 19, &H827E6C: SetPixelV lHDC, tmpw + 25, tmph + 19, &H6A665B: SetPixelV lHDC, tmpw + 26, tmph + 19, &H625F5A: SetPixelV lHDC, tmpw + 27, tmph + 19, &H8B8C8C: SetPixelV lHDC, tmpw + 28, tmph + 19, &HCDCDCD: SetPixelV lHDC, tmpw + 29, tmph + 19, &HE8E8E8: SetPixelV lHDC, tmpw + 30, tmph + 19, &HFAFAFA:
  SetPixelV lHDC, tmpw + 17, tmph + 20, &H5A563E: SetPixelV lHDC, tmpw + 18, tmph + 20, &H59543D: SetPixelV lHDC, tmpw + 19, tmph + 20, &H58513A: SetPixelV lHDC, tmpw + 20, tmph + 20, &H554F38: SetPixelV lHDC, tmpw + 21, tmph + 20, &H59513B: SetPixelV lHDC, tmpw + 22, tmph + 20, &H58513E: SetPixelV lHDC, tmpw + 23, tmph + 20, &H646053: SetPixelV lHDC, tmpw + 24, tmph + 20, &H7B7973: SetPixelV lHDC, tmpw + 25, tmph + 20, &HA2A19F: SetPixelV lHDC, tmpw + 26, tmph + 20, &HC5C5C5: SetPixelV lHDC, tmpw + 27, tmph + 20, &HDADADA: SetPixelV lHDC, tmpw + 28, tmph + 20, &HEDEDED: SetPixelV lHDC, tmpw + 29, tmph + 20, &HFAFAFA:
  SetPixelV lHDC, tmpw + 17, tmph + 21, &HC5C5C5: SetPixelV lHDC, tmpw + 18, tmph + 21, &HC5C5C5: SetPixelV lHDC, tmpw + 19, tmph + 21, &HC6C6C6: SetPixelV lHDC, tmpw + 20, tmph + 21, &HC6C6C6: SetPixelV lHDC, tmpw + 21, tmph + 21, &HC6C6C6: SetPixelV lHDC, tmpw + 22, tmph + 21, &HC9C9C9: SetPixelV lHDC, tmpw + 23, tmph + 21, &HCECECE: SetPixelV lHDC, tmpw + 24, tmph + 21, &HD7D7D7: SetPixelV lHDC, tmpw + 25, tmph + 21, &HE1E1E1: SetPixelV lHDC, tmpw + 26, tmph + 21, &HECECEC: SetPixelV lHDC, tmpw + 27, tmph + 21, &HF6F6F6: SetPixelV lHDC, tmpw + 28, tmph + 21, &HFDFDFD:
  SetPixelV lHDC, tmpw + 17, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 18, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 19, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 20, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 21, tmph + 22, &HECECEC: SetPixelV lHDC, tmpw + 22, tmph + 22, &HEDEDED: SetPixelV lHDC, tmpw + 23, tmph + 22, &HF0F0F0: SetPixelV lHDC, tmpw + 24, tmph + 22, &HF4F4F4: SetPixelV lHDC, tmpw + 25, tmph + 22, &HFAFAFA: SetPixelV lHDC, tmpw + 26, tmph + 22, &HFDFDFD:
  tmph = 11:     tmph1 = lh - 10:     tmpw = lw - 34
  'Generar lineas intermedias
  APILine 0, tmph, 0, tmph1, &HF7F7F7: APILine 1, tmph, 1, tmph1, &HA99D9B: APILine 2, tmph, 2, tmph1, &H632D17: APILine 3, tmph, 3, tmph1, &HA65A1D: APILine 4, tmph, 4, tmph1, &HB96C2E
  APILine 5, tmph, 5, tmph1, &HBC7738: APILine 6, tmph, 6, tmph1, &HC18242: APILine 7, tmph, 7, tmph1, &HC2894E: APILine 8, tmph, 8, tmph1, &HC18A52: APILine 9, tmph, 9, tmph1, &HC59157
  APILine 10, tmph, 10, tmph1, &HC59159: APILine 11, tmph, 11, tmph1, &HCC9863: APILine 12, tmph, 12, tmph1, &HCC9665: APILine 13, tmph, 13, tmph1, &HCB9767: APILine 14, tmph, 14, tmph1, &HC99565
  APILine 15, tmph, 15, tmph1, &HCC9A6A: APILine 16, tmph, 16, tmph1, &HCC9A6A: APILine 17, tmph, 17, tmph1, &HCC9B6A: APILine tmpw + 17, tmph, tmpw + 17, tmph1, &HCC9B6A: APILine tmpw + 18, tmph, tmpw + 18, tmph1, &HCC9B6A:
  APILine tmpw + 19, tmph, tmpw + 19, tmph1, &HCA9864: APILine tmpw + 20, tmph, tmpw + 20, tmph1, &HC99763: APILine tmpw + 21, tmph, tmpw + 21, tmph1, &HC99763: APILine tmpw + 22, tmph, tmpw + 22, tmph1, &HCA9562: APILine tmpw + 23, tmph, tmpw + 23, tmph1, &HC9945D
  APILine tmpw + 24, tmph, tmpw + 24, tmph1, &HC68E57: APILine tmpw + 25, tmph, tmpw + 25, tmph1, &HC48C55: APILine tmpw + 26, tmph, tmpw + 26, tmph1, &HC3884E: APILine tmpw + 27, tmph, tmpw + 27, tmph1, &HBE823E: APILine tmpw + 28, tmph, tmpw + 28, tmph1, &HB97634
  APILine tmpw + 29, tmph, tmpw + 29, tmph1, &HBD6D2D: APILine tmpw + 30, tmph, tmpw + 30, tmph1, &HA8581C: APILine tmpw + 31, tmph, tmpw + 31, tmph1, &H6D3319: APILine tmpw + 32, tmph, tmpw + 32, tmph1, &HA49794: APILine tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6
  'Lineas verticales
  APILine 17, 0, lw - 17, 0, &H3C090A
  APILine 17, 1, lw - 17, 1, &HDEC7BE
  APILine 17, 2, lw - 17, 2, &HD3BCB2
  APILine 17, 3, lw - 17, 3, &HD4B39A
  APILine 17, 4, lw - 17, 4, &HCCAC92
  APILine 17, 5, lw - 17, 5, &HCFA88A
  APILine 17, 6, lw - 17, 6, &HCCA587
  APILine 17, 7, lw - 17, 7, &HD2A77D
  APILine 17, 8, lw - 17, 8, &HB98E62
  APILine 17, 9, lw - 17, 9, &HC69464
  APILine 17, 10, lw - 17, 10, &HCC9B6A
  APILine 17, 11, lw - 17, 11, &HD0A175
  tmph = lh - 22
  APILine 17, tmph + 11, lw - 17, tmph + 11, &HD0A175
  APILine 17, tmph + 12, lw - 17, tmph + 12, &HDAAA7F
  APILine 17, tmph + 13, lw - 17, tmph + 13, &HDCB087
  APILine 17, tmph + 14, lw - 17, tmph + 14, &HE4B88E
  APILine 17, tmph + 15, lw - 17, tmph + 15, &HE1BF93
  APILine 17, tmph + 16, lw - 17, tmph + 16, &HE7C699
  APILine 17, tmph + 17, lw - 17, tmph + 17, &HE2CC9D
  APILine 17, tmph + 18, lw - 17, tmph + 18, &HE9D5A6
  APILine 17, tmph + 19, lw - 17, tmph + 19, &HD7D2BA
  APILine 17, tmph + 20, lw - 17, tmph + 20, &H5A563E
  APILine 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
  APILine 17, tmph + 22, lw - 17, tmph + 22, &HECECEC
  Exit Sub

DrawMacOSXButtonPressed_Error:
End Sub

Private Sub DrawPlastikButton(iState As isState)
  On Error GoTo DrawPlastikButton_Error

  Dim tmpcolor As Long

  'If Ambient.DisplayAsDefault Then iState = stateDefaulted
  Select Case iState
    Case statenormal, stateDefaulted
      tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H8)
      UserControl.BackColor = tmpcolor
      DrawVGradient OffsetColor(tmpcolor, &HF), OffsetColor(tmpcolor, -&HF), 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3
      DrawVGradient OffsetColor(tmpcolor, &H15), OffsetColor(tmpcolor, -&H15), 1, 2, 2, UserControl.ScaleHeight - 5
      DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H20), UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5
      tmpcolor = OffsetColor(tmpcolor, -&H60)
      APILine 2, 0, UserControl.ScaleWidth - 2, 0, tmpcolor
      APILine 0, 2, 0, UserControl.ScaleHeight - 2, tmpcolor
      APILine 2, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      APILine UserControl.ScaleWidth - 1, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight, tmpcolor
      SetPixelV UserControl.hdc, 1, 1, tmpcolor: SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 2, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, tmpcolor
      'Border Pixels
      tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H15)
      SetPixelV UserControl.hdc, 1, 0, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 0, tmpcolor
      SetPixelV UserControl.hdc, 0, 1, tmpcolor: SetPixelV UserControl.hdc, 0, UserControl.ScaleHeight - 2, tmpcolor
      SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 2, tmpcolor
      APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H15)
      APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H25)

      If iState = stateDefaulted Or m_bFocused Then
        tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
        APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(tmpcolor, &H15)
        APILine 1, 2, UserControl.ScaleWidth - 1, 2, OffsetColor(tmpcolor, &H15)
        APILine 1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H5)
        APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(tmpcolor, -&H15)
      End If

      Exit Sub

    Case stateHot
      tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H18)
      UserControl.BackColor = tmpcolor
      DrawVGradient OffsetColor(tmpcolor, &H10), OffsetColor(tmpcolor, -&H10), 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3
      DrawVGradient OffsetColor(tmpcolor, &H15), OffsetColor(tmpcolor, -&H15), 1, 2, 2, UserControl.ScaleHeight - 5
      DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H20), UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5
      tmpcolor = OffsetColor(tmpcolor, -&H60)
      APILine 2, 0, UserControl.ScaleWidth - 2, 0, tmpcolor
      APILine 0, 2, 0, UserControl.ScaleHeight - 2, tmpcolor
      APILine 2, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      APILine UserControl.ScaleWidth - 1, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight, tmpcolor
      SetPixelV UserControl.hdc, 1, 1, tmpcolor: SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 2, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, tmpcolor
      'Border Pixels
      tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H15)
      SetPixelV UserControl.hdc, 1, 0, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 0, tmpcolor
      SetPixelV UserControl.hdc, 0, 1, tmpcolor: SetPixelV UserControl.hdc, 0, UserControl.ScaleHeight - 2, tmpcolor
      SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 2, tmpcolor
      APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(GetSysColor(COLOR_BTNFACE), &H15)
      APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(GetSysColor(COLOR_BTNFACE), -&H10)
      tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
      APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(tmpcolor, &H15)
      APILine 1, 2, UserControl.ScaleWidth - 1, 2, OffsetColor(tmpcolor, &H15)
      APILine 1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H5)
      APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(tmpcolor, -&H15)
      Exit Sub

    Case statePressed
      tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H9)
      DrawVGradient OffsetColor(tmpcolor, -&HF), tmpcolor, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
      DrawVGradient OffsetColor(tmpcolor, -&H15), OffsetColor(tmpcolor, &H15), UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5
      DrawVGradient OffsetColor(tmpcolor, -&H20), OffsetColor(tmpcolor, -&H5), 1, 2, 2, UserControl.ScaleHeight - 5
      tmpcolor = OffsetColor(tmpcolor, -&H60)
      APILine 2, 0, UserControl.ScaleWidth - 2, 0, tmpcolor
      APILine 0, 2, 0, UserControl.ScaleHeight - 2, tmpcolor
      APILine 2, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      APILine UserControl.ScaleWidth - 1, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight, tmpcolor
      SetPixelV UserControl.hdc, 1, 1, tmpcolor: SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 2, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, tmpcolor
      'Border Pixels
      tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H15)
      SetPixelV UserControl.hdc, 1, 0, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 0, tmpcolor
      SetPixelV UserControl.hdc, 0, 1, tmpcolor: SetPixelV UserControl.hdc, 0, UserControl.ScaleHeight - 2, tmpcolor
      SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 2, tmpcolor
      APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H8)
      APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H15)
      Exit Sub

    Case statedisabled
      tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H12)
      UserControl.BackColor = tmpcolor
      DrawVGradient OffsetColor(tmpcolor, &HF), OffsetColor(tmpcolor, -&HF), 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3
      DrawVGradient OffsetColor(tmpcolor, &H15), OffsetColor(tmpcolor, -&H15), 1, 2, 2, UserControl.ScaleHeight - 5
      DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H20), UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 5
      tmpcolor = OffsetColor(tmpcolor, -&H60)
      APILine 2, 0, UserControl.ScaleWidth - 2, 0, tmpcolor
      APILine 0, 2, 0, UserControl.ScaleHeight - 2, tmpcolor
      APILine 2, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      APILine UserControl.ScaleWidth - 1, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight, tmpcolor
      SetPixelV UserControl.hdc, 1, 1, tmpcolor: SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 2, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, tmpcolor
      'Border Pixels
      tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H5)
      SetPixelV UserControl.hdc, 1, 0, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, 0, tmpcolor
      SetPixelV UserControl.hdc, 0, 1, tmpcolor: SetPixelV UserControl.hdc, 0, UserControl.ScaleHeight - 2, tmpcolor
      SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1, tmpcolor
      SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 2, tmpcolor
      APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H15)
      APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&H25)
      'If iState = stateDefaulted Or m_bFocused Then
      '    tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
      '    APILine 2, 1, UserControl.ScaleWidth - 2, 1, OffsetColor(tmpcolor, &H15)
      '    APILine 1, 2, UserControl.ScaleWidth - 1, 2, OffsetColor(tmpcolor, &H15)
      '    APILine 1, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H5)
      '    APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, OffsetColor(tmpcolor, -&H15)
      'End If
      Exit Sub

  End Select

  Exit Sub

DrawPlastikButton_Error:
End Sub

Private Sub DrawGalaxyButton(iState As isState)
  On Error GoTo DrawGalaxyButton_Error

  Dim tmpcolor As Long

  If iState = statenormal Then
    tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
  Else
    tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, OffsetColor(GetSysColor(COLOR_BTNFACE), &HF))
  End If

  UserControl.BackColor = tmpcolor

  If iState = statePressed Then
    DrawVGradient OffsetColor(tmpcolor, -&HF), OffsetColor(tmpcolor, &HF), 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 6
    APILine 2, 1, UserControl.ScaleWidth - 3, 1, tmpcolor
    APILine 1, 2, 1, UserControl.ScaleHeight - 3, tmpcolor
    APILine 2, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, &H60)
    APILine UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, &H60)
  Else
    DrawVGradient OffsetColor(tmpcolor, &HF), OffsetColor(tmpcolor, -&HF), 2, 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 6
    APILine 2, 1, UserControl.ScaleWidth - 3, 1, OffsetColor(tmpcolor, &H60)
    APILine 1, 2, 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, &H60)
    APILine 2, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, tmpcolor
    APILine UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, tmpcolor
  End If

  tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))

  If iState <> statePressed Then
    tmpcolor = IIf(m_bFocused, OffsetColor(tmpcolor, -&H60), OffsetColor(tmpcolor, -&H30))
  Else
    tmpcolor = OffsetColor(tmpcolor, -&H30)
  End If

  APILine 2, 0, UserControl.ScaleWidth - 3, 0, tmpcolor
  APILine 0, 2, 0, UserControl.ScaleHeight - 3, tmpcolor
  APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 2, tmpcolor
  APILine UserControl.ScaleWidth - 2, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 3, tmpcolor
  SetPixelV UserControl.hdc, 1, 1, tmpcolor: SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 3, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 3, 1, tmpcolor: SetPixelV UserControl.hdc, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, tmpcolor
  tmpcolor = OffsetColor(tmpcolor, &H15)
  APILine 3, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 1, tmpcolor
  APILine UserControl.ScaleWidth - 1, 3, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 4, tmpcolor
  APILine UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 1, UserControl.ScaleWidth, UserControl.ScaleHeight - 5, tmpcolor
  APILine UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 1, UserControl.ScaleWidth, UserControl.ScaleHeight - 4, OffsetColor(tmpcolor, &HF)
  Exit Sub

DrawGalaxyButton_Error:
End Sub

Private Sub DrawKeramikButton(iState As isState)
  On Error GoTo DrawKeramikButton_Error

  Dim tmpcolor As Long

  If m_iState = statenormal Then
    tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), -&HF)
  ElseIf m_iState = stateHot Then
    tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE)), &H1)
  ElseIf m_iState = statePressed Then
    tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE)), &H18)
  Else
    tmpcolor = OffsetColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &HF)
  End If

  UserControl.BackColor = tmpcolor
  DrawVGradient OffsetColor(tmpcolor, &H20), OffsetColor(tmpcolor, -&H20), 5, 2, UserControl.ScaleWidth - 5, UserControl.ScaleHeight - 6
  'Left
  DrawVGradient OffsetColor(tmpcolor, &H80), OffsetColor(tmpcolor, -&H80), 0, 0, 1, UserControl.ScaleHeight - 4
  DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H25), 2, 4, 3, UserControl.ScaleHeight / 2 - 7
  DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H35), 3, 3, 4, UserControl.ScaleHeight / 2 - 7
  DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H25), 4, 2, 5, UserControl.ScaleHeight / 2 - 7
  DrawVGradient OffsetColor(tmpcolor, &H25), OffsetColor(tmpcolor, &H5), 5, 2, 6, UserControl.ScaleHeight / 2 - 4
  DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H20), 2, UserControl.ScaleHeight / 2 - 2, 3, UserControl.ScaleHeight / 2 - 5
  DrawVGradient OffsetColor(tmpcolor, -&H35), OffsetColor(tmpcolor, -&H20), 3, UserControl.ScaleHeight / 2 - 3, 4, UserControl.ScaleHeight / 2 - 3
  DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H20), 4, UserControl.ScaleHeight / 2 - 4, 5, UserControl.ScaleHeight / 2 - 1
  DrawVGradient OffsetColor(tmpcolor, &H5), OffsetColor(tmpcolor, -&H5), 5, UserControl.ScaleHeight / 2 - 4, 6, UserControl.ScaleHeight / 2 - 3
  DrawHGradient OffsetColor(tmpcolor, -&H35), OffsetColor(tmpcolor, -&H20), 5, UserControl.ScaleHeight - 5, 12, UserControl.ScaleHeight - 6
  DrawHGradient OffsetColor(tmpcolor, -&H38), OffsetColor(tmpcolor, -&H20), 4, UserControl.ScaleHeight - 6, 7, UserControl.ScaleHeight - 7
  DrawHGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H20), 3, UserControl.ScaleHeight - 7, 5, UserControl.ScaleHeight - 8
  'Right
  DrawVGradient OffsetColor(tmpcolor, &H80), OffsetColor(tmpcolor, -&H80), UserControl.ScaleWidth - 2, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 4
  DrawVGradient OffsetColor(tmpcolor, &H80), OffsetColor(tmpcolor, -&H30), UserControl.ScaleWidth - 1, 0, UserControl.ScaleWidth, UserControl.ScaleHeight / 2
  DrawVGradient OffsetColor(tmpcolor, -&H40), OffsetColor(tmpcolor, -&H8), UserControl.ScaleWidth - 1, UserControl.ScaleHeight / 2, UserControl.ScaleWidth, UserControl.ScaleHeight / 2 - 4
  DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H25), UserControl.ScaleWidth - 4, 4, UserControl.ScaleWidth - 3, UserControl.ScaleHeight / 2 - 5
  DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H30), UserControl.ScaleWidth - 5, 3, UserControl.ScaleWidth - 4, UserControl.ScaleHeight / 2 - 4
  DrawVGradient tmpcolor, OffsetColor(tmpcolor, -&H25), UserControl.ScaleWidth - 6, 2, UserControl.ScaleWidth - 5, UserControl.ScaleHeight / 2 - 3
  DrawVGradient OffsetColor(tmpcolor, &H25), OffsetColor(tmpcolor, &H5), UserControl.ScaleWidth - 7, UserControl.ScaleWidth - 6, 6, UserControl.ScaleHeight / 2 - 4
  DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H15), UserControl.ScaleWidth - 4, UserControl.ScaleHeight / 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight / 2 - 5
  DrawVGradient OffsetColor(tmpcolor, -&H30), OffsetColor(tmpcolor, -&H25), UserControl.ScaleWidth - 5, UserControl.ScaleHeight / 2, UserControl.ScaleWidth - 4, UserControl.ScaleHeight / 2 - 5
  DrawVGradient OffsetColor(tmpcolor, -&H25), OffsetColor(tmpcolor, -&H35), UserControl.ScaleWidth - 6, UserControl.ScaleHeight / 2, UserControl.ScaleWidth - 5, UserControl.ScaleHeight / 2 - 4
  DrawVGradient OffsetColor(tmpcolor, -&H5), OffsetColor(tmpcolor, -&H25), UserControl.ScaleWidth - 7, UserControl.ScaleHeight / 2, UserControl.ScaleWidth - 6, UserControl.ScaleHeight / 2 - 4
  DrawHGradient OffsetColor(tmpcolor, -&H20), OffsetColor(tmpcolor, -&H35), UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 4, UserControl.ScaleWidth - 7, UserControl.ScaleHeight - 3
  'top
  APILine 3, 0, UserControl.ScaleWidth - 3, 0, OffsetColor(tmpcolor, &H30)
  APILine 1, 1, UserControl.ScaleWidth - 1, 1, OffsetColor(tmpcolor, &H30)
  APILine 5, 1, UserControl.ScaleWidth - 5, 1, tmpcolor 'OffsetColor(tmpcolor, &H30)
  DrawHGradient OffsetColor(tmpcolor, &H20), tmpcolor, UserControl.ScaleWidth - 11, 2, UserControl.ScaleWidth - 4, 3
  DrawHGradient OffsetColor(tmpcolor, &H20), tmpcolor, UserControl.ScaleWidth - 10, 3, UserControl.ScaleWidth - 5, 4
  DrawHGradient OffsetColor(tmpcolor, &H20), tmpcolor, UserControl.ScaleWidth - 9, 4, UserControl.ScaleWidth - 6, 5
  APILine 6, 3, 7, 3, OffsetColor(tmpcolor, &H80)
  'bottom
  APILine 3, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 1, OffsetColor(tmpcolor, -&H10)
  APILine 2, UserControl.ScaleHeight - 2, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 2, OffsetColor(tmpcolor, -&H80)
  APILine 7, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 7, UserControl.ScaleHeight - 3, tmpcolor
  SetPixelV UserControl.hdc, 1, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H70): SetPixelV UserControl.hdc, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, OffsetColor(tmpcolor, -&H70)
  APILine UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 4, OffsetColor(tmpcolor, -&H15)
  Exit Sub

DrawKeramikButton_Error:
End Sub

'Important. If not included, tooltips don't change when you try to set the toltip text
Private Sub RemoveToolTip()
  On Error GoTo RemoveToolTip_Error

  Dim lR As Long
  UserControl.Extender.ToolTipText = m_sToolTipText

  If m_lttHwnd <> 0 Then
    '      With ttip
    '         .lSize = Len(ttip)
    '         .lHwnd = UserControl.hwnd
    '      End With
    lR = SendMessage(ttip.lHwnd, TTM_DELTOOLA, 0, ttip)
    DestroyWindow m_lttHwnd
  End If

  Exit Sub

RemoveToolTip_Error:
End Sub

Private Sub CreateToolTip()
  On Error GoTo CreateToolTip_Error

  Dim lpRect As RECT
  Dim lWinStyle As Long
  'RemoveToolTip
  m_sToolTipText = UserControl.Extender.ToolTipText
  UserControl.Extender.ToolTipText = ""
  ttip.lpStr = m_sToolTipText
  lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
  
  If m_bEnabled = False Or m_bVisible = False Then Exit Sub
  
  ''create baloon style if desired
  If m_lToolTipType = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
  m_lttHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hWnd, 0&, App.hInstance, 0&)
  ''make our tooltip window a topmost window
  'This has been causing troubles :( So, It's temporally comented
  'SetWindowPos m_lttHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
  ''get the rect of the parent control
  GetClientRect UserControl.hWnd, lpRect

  ''now set our tooltip info structure
  With ttip

    ''if we want it centered, then set that flag
    If m_lttCentered Then
      .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
    Else
      .lFlags = TTF_SUBCLASS
    End If

    ''set the hwnd prop to our parent control's hwnd
    .lHwnd = UserControl.hWnd
    .lId = 0
    .hInstance = App.hInstance
    '.lpstr = ALREADY SET
    .lpRect = lpRect
  End With

  ''add the tooltip structure
  SendMessage m_lttHwnd, TTM_ADDTOOLA, 0&, ttip

  ''if we want a title or we want an icon
  If m_sTooltiptitle <> vbNullString Or m_lToolTipIcon <> TTNoIcon Then
    SendMessage m_lttHwnd, TTM_SETTITLE, CLng(m_lToolTipIcon), ByVal m_sTooltiptitle
  End If

  If m_lttForeColor <> Empty Then
    SendMessage m_lttHwnd, TTM_SETTIPTEXTCOLOR, TranslateColor(m_lttForeColor), 0&
  End If

  If m_lttBackColor <> Empty Then
    SendMessage m_lttHwnd, TTM_SETTIPBKCOLOR, TranslateColor(m_lttBackColor), 0&
  End If

  Exit Sub

CreateToolTip_Error:
End Sub

Private Sub m_About_Click()
  On Error GoTo m_About_Click_Error

  m_About.Visible = False
  Exit Sub

m_About_Click_Error:
End Sub

Private Sub m_About_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
  On Error GoTo m_About_MouseDown_Error

  Dim tmpRect As RECT
  Dim tmpcolor As Long

  With m_About
    'Draw button
    SetRect tmpRect, 290, 80, 360, 26
    'tmpcolor = OffsetColor(GetSysColor(COLOR_BTNFACE), &HF)
    tmpcolor = GetSysColor(COLOR_BTNFACE)
    DrawVGradientEx m_About.hdc, OffsetColor(tmpcolor, -&HF), OffsetColor(tmpcolor, &HF), tmpRect.Left, tmpRect.Top, tmpRect.Right, tmpRect.Bottom
    tmpcolor = GetSysColor(COLOR_BTNSHADOW)
    APILineEx .hdc, tmpRect.Left + 2, tmpRect.Top, tmpRect.Right - 2, tmpRect.Top, tmpcolor
    APILineEx .hdc, tmpRect.Left + 2, tmpRect.Top, tmpRect.Left, tmpRect.Top + 2, tmpcolor
    APILineEx .hdc, tmpRect.Right - 2, tmpRect.Top, tmpRect.Right, tmpRect.Top + 2, tmpcolor
    APILineEx .hdc, tmpRect.Left, tmpRect.Top + 2, tmpRect.Left, tmpRect.Top + tmpRect.Bottom - 2, tmpcolor
    APILineEx .hdc, tmpRect.Right, tmpRect.Top + 2, tmpRect.Right, tmpRect.Top + tmpRect.Bottom, tmpcolor
    APILineEx .hdc, tmpRect.Left, tmpRect.Top + tmpRect.Bottom - 2, tmpRect.Left + 2, tmpRect.Top + tmpRect.Bottom, tmpcolor
    APILineEx .hdc, tmpRect.Right, tmpRect.Top + tmpRect.Bottom - 2, tmpRect.Right - 3, tmpRect.Top + tmpRect.Bottom + 1, tmpcolor
    APILineEx .hdc, tmpRect.Left + 2, tmpRect.Top + tmpRect.Bottom, tmpRect.Right - 2, tmpRect.Top + tmpRect.Bottom, tmpcolor
    tmpcolor = GetSysColor(COLOR_BTNFACE)
    SetPixelV .hdc, tmpRect.Left, tmpRect.Top, tmpcolor
    SetPixelV .hdc, tmpRect.Left + 1, tmpRect.Top, tmpcolor
    SetPixelV .hdc, tmpRect.Left, tmpRect.Top + 1, tmpcolor
    SetPixelV .hdc, tmpRect.Right, tmpRect.Top, tmpcolor
    SetPixelV .hdc, tmpRect.Right - 1, tmpRect.Top, tmpcolor
    SetPixelV .hdc, tmpRect.Right, tmpRect.Top + 1, tmpcolor
    SetPixelV .hdc, tmpRect.Left, tmpRect.Top + tmpRect.Bottom, tmpcolor
    SetPixelV .hdc, tmpRect.Left + 1, tmpRect.Top + tmpRect.Bottom, tmpcolor
    SetPixelV .hdc, tmpRect.Left, tmpRect.Top + tmpRect.Bottom - 1, tmpcolor
    SetPixelV .hdc, tmpRect.Right, tmpRect.Top + tmpRect.Bottom, tmpcolor
    SetPixelV .hdc, tmpRect.Right - 1, tmpRect.Top + tmpRect.Bottom, tmpcolor
    SetPixelV .hdc, tmpRect.Right, tmpRect.Top + tmpRect.Bottom - 1, tmpcolor
    SetRect tmpRect, 290, 80, 360, 106
    .FontSize = 8
    .ForeColor = vbBlue
    .FontUnderline = True
    DrawText .hdc, "Close", -1, tmpRect, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE
  End With

  Exit Sub

m_About_MouseDown_Error:
End Sub

Private Sub m_About_Paint()
  'Draw the About content
  On Error GoTo m_About_Paint_Error

  Dim lwformat As Long
  Dim tmpRect As RECT
  Dim tmpcolor As Long
  lwformat = DT_VCENTER Or DT_LEFT Or DT_SINGLELINE

  With m_About
    .ForeColor = GetSysColor(COLOR_BTNTEXT)
    .FontUnderline = False
    .FontSize = 18
    SetRect tmpRect, 20, 10, 220, 40
    DrawText .hdc, "isButton", -1, tmpRect, lwformat
    .FontSize = 10
    SetRect tmpRect, 160, 70, 300, 20
    DrawText .hdc, "Version " & strCurrentVersion, -1, tmpRect, lwformat
    .FontBold = True
    SetRect tmpRect, 20, 110, 250, 20
    DrawText .hdc, "By Fred.cpp", -1, tmpRect, lwformat
    .FontBold = False
    SetRect tmpRect, 20, 140, 250, 20
    DrawText .hdc, "http://mx.geocities.com/fred_cpp/", -1, tmpRect, lwformat
    'Draw button
    SetRect tmpRect, 290, 80, 360, 26
    'tmpcolor = OffsetColor(GetSysColor(COLOR_BTNFACE), &HF)
    tmpcolor = GetSysColor(COLOR_BTNFACE)
    DrawVGradientEx m_About.hdc, OffsetColor(tmpcolor, &HF), OffsetColor(tmpcolor, -&HF), tmpRect.Left, tmpRect.Top, tmpRect.Right, tmpRect.Bottom
    tmpcolor = GetSysColor(COLOR_BTNSHADOW)
    APILineEx .hdc, tmpRect.Left + 2, tmpRect.Top, tmpRect.Right - 2, tmpRect.Top, tmpcolor
    APILineEx .hdc, tmpRect.Left + 2, tmpRect.Top, tmpRect.Left, tmpRect.Top + 2, tmpcolor
    APILineEx .hdc, tmpRect.Right - 2, tmpRect.Top, tmpRect.Right, tmpRect.Top + 2, tmpcolor
    APILineEx .hdc, tmpRect.Left, tmpRect.Top + 2, tmpRect.Left, tmpRect.Top + tmpRect.Bottom - 2, tmpcolor
    APILineEx .hdc, tmpRect.Right, tmpRect.Top + 2, tmpRect.Right, tmpRect.Top + tmpRect.Bottom, tmpcolor
    APILineEx .hdc, tmpRect.Left, tmpRect.Top + tmpRect.Bottom - 2, tmpRect.Left + 2, tmpRect.Top + tmpRect.Bottom, tmpcolor
    APILineEx .hdc, tmpRect.Right, tmpRect.Top + tmpRect.Bottom - 2, tmpRect.Right - 3, tmpRect.Top + tmpRect.Bottom + 1, tmpcolor
    APILineEx .hdc, tmpRect.Left + 2, tmpRect.Top + tmpRect.Bottom, tmpRect.Right - 2, tmpRect.Top + tmpRect.Bottom, tmpcolor
    tmpcolor = GetSysColor(COLOR_BTNFACE)
    SetPixelV .hdc, tmpRect.Left, tmpRect.Top, tmpcolor
    SetPixelV .hdc, tmpRect.Left + 1, tmpRect.Top, tmpcolor
    SetPixelV .hdc, tmpRect.Left, tmpRect.Top + 1, tmpcolor
    SetPixelV .hdc, tmpRect.Right, tmpRect.Top, tmpcolor
    SetPixelV .hdc, tmpRect.Right - 1, tmpRect.Top, tmpcolor
    SetPixelV .hdc, tmpRect.Right, tmpRect.Top + 1, tmpcolor
    SetPixelV .hdc, tmpRect.Left, tmpRect.Top + tmpRect.Bottom, tmpcolor
    SetPixelV .hdc, tmpRect.Left + 1, tmpRect.Top + tmpRect.Bottom, tmpcolor
    SetPixelV .hdc, tmpRect.Left, tmpRect.Top + tmpRect.Bottom - 1, tmpcolor
    SetPixelV .hdc, tmpRect.Right, tmpRect.Top + tmpRect.Bottom, tmpcolor
    SetPixelV .hdc, tmpRect.Right - 1, tmpRect.Top + tmpRect.Bottom, tmpcolor
    SetPixelV .hdc, tmpRect.Right, tmpRect.Top + tmpRect.Bottom - 1, tmpcolor
    SetRect tmpRect, 290, 80, 360, 106
    .FontSize = 8
    .ForeColor = vbBlue
    .FontUnderline = True
    DrawText .hdc, "Close", -1, tmpRect, DT_VCENTER Or DT_CENTER Or DT_SINGLELINE
  End With

  Exit Sub

m_About_Paint_Error:
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  'When action called by Accesskey
  On Error GoTo UserControl_AccessKeyPress_Error

  lPrevButton = vbLeftButton
  UserControl_Click
  m_iState = statenormal
  RaiseEvent Click
  Exit Sub

UserControl_AccessKeyPress_Error:
End Sub

'Manage default and Cancel events
'Looks like each time the control get's the focus, the Default
' property is also Set, It's kinda annoying I still am Trying
' to figure out why and how can I Implement Default And Cancel
' properties
Private Sub UserControl_AmbientChanged(PropertyName As String)
  On Error GoTo UserControl_AmbientChanged_Error

  Select Case PropertyName
    Case "DisplayAsDefault"
      Refresh
  End Select

  Exit Sub

UserControl_AmbientChanged_Error:
End Sub

Private Sub UserControl_Click()
  On Error GoTo UserControl_Click_Error

  If lPrevButton = vbLeftButton Then
    If m_ButtonType = isbCheckBox Then
      m_Value = Not m_Value
    End If

    m_iState = stateHot
    Refresh
    
  End If

  Exit Sub

UserControl_Click_Error:
End Sub

Private Sub UserControl_DblClick()
  On Error GoTo UserControl_DblClick_Error

  If lPrevButton = vbLeftButton Then
    UserControl_MouseDown 1, 0, 1, 1
  End If

  Exit Sub

UserControl_DblClick_Error:
End Sub

Private Sub UserControl_EnterFocus()
  On Error GoTo UserControl_EnterFocus_Error

  m_bFocused = True
  Refresh
  Exit Sub

UserControl_EnterFocus_Error:
End Sub

Private Sub UserControl_ExitFocus()
  On Error GoTo UserControl_ExitFocus_Error

  m_bFocused = False
  Refresh
  Exit Sub

UserControl_ExitFocus_Error:
End Sub

Private Sub UserControl_GotFocus()
  On Error GoTo UserControl_GotFocus_Error

  m_bFocused = True
  Refresh
  Exit Sub

UserControl_GotFocus_Error:
End Sub

Private Sub UserControl_Hide()
  On Error GoTo UserControl_Hide_Error

  m_bVisible = False
  UserControl.Extender.ToolTipText = m_sToolTipText
  Exit Sub

UserControl_Hide_Error:
End Sub

Private Sub UserControl_InitProperties()
  On Error GoTo UserControl_InitProperties_Error

  m_iStyle = 0
  m_sCaption = UserControl.Extender.Name
  m_IconSize = 16
  Set m_Icon = LoadPicture
  lwFontAlign = DT_CENTER Or DT_WORDBREAK 'DT_VCENTER Or DT_CENTER
  m_bEnabled = True
  m_bShowFocus = False
  m_bUseCustomColors = False
  m_lBackColor = TranslateColor(vbButtonFace)
  m_lHighlightColor = TranslateColor(vbHighlight)
  m_lFontColor = TranslateColor(vbButtonText)
  m_lFontHighlightColor = TranslateColor(vbButtonText)
  lPrevStyle = GetWindowLong(m_About.hWnd, GWL_STYLE)
  m_lToolTipType = TTBalloon
  m_CaptionAlign = isbCenter
  m_IconAlign = isbleft
  iStyleIconOffset = 4
  Exit Sub

UserControl_InitProperties_Error:
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
  On Error GoTo UserControl_KeyDown_Error

  Select Case KeyCode
    Case vbKeySpace
      m_iState = statePressed
      Refresh
    Case vbKeyRight, vbKeyDown
      SendKeys "{TAB}"
    Case vbKeyLeft, vbKeyUp
      SendKeys "+{TAB}"
    Case vbKeyTab
      RaiseEvent MouseLeave
  End Select

  Exit Sub

UserControl_KeyDown_Error:
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        UserControl_Click
        RaiseEvent Click
        m_iState = statenormal
        Refresh
    End If
End Sub

Private Sub UserControl_LostFocus()
  On Error GoTo UserControl_LostFocus_Error

  m_bFocused = False
  Refresh
  Exit Sub

UserControl_LostFocus_Error:
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
  On Error GoTo UserControl_MouseDown_Error
  
  If Me.Enabled = False Then Exit Sub

  If Button = vbLeftButton Then
    m_iState = statePressed
    Refresh
  End If

  Exit Sub

UserControl_MouseDown_Error:
End Sub

' Description: Refresh the control
Public Sub Refresh()
  Dim tmpcolor As Long
  Dim lTransColor As Long
  Dim lcurrpix As Long
  
  UserControl.Cls
  iStyleIconOffset = 3

  If Not m_bVisible Then Exit Sub
  If Not UserControl.Ambient.UserMode And m_iStyle <> isbWindowsXP Then m_iState = stateHot
  If m_ButtonType = isbCheckBox Then
    If m_Value Then m_iState = statePressed
  End If

  If Not m_bEnabled Then
    m_iState = statedisabled
    UserControl.BackColor = GetSysColor(COLOR_BTNFACE)
  End If

  Select Case m_iStyle
    Case isbNormal

      'Classic Style (Win98)
      If m_iState = statenormal Then
        tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
      Else
        tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))
      End If

      UserControl.BackColor = tmpcolor
      DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, IIf(m_iState = statePressed, EDGE_SUNKEN, EDGE_RAISED)

      'If Ambient.DisplayAsDefault Then
      If (m_bFocused And m_bShowFocus) Or (Ambient.DisplayAsDefault And Not m_bFocused) Then
        APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, vbBlack
      End If

      'End If
    Case isbSoft

      'Soft Style (I don't know where does It come, But I've seen this before)
      If m_iState = statenormal Or m_iState = statedisabled Then
        tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
      Else
        tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))
      End If

      UserControl.BackColor = tmpcolor

      Select Case m_iState
        Case statenormal, stateHot, stateDefaulted
          DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_RAISEDINNER
        Case statePressed
          DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER
        Case statedisabled
          APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
      End Select

      'If Ambient.DisplayAsDefault Then
      '    APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, vbBlack
      'End If
    Case isbFlat

      'Flat Style (Office 2000 like)
      If m_iState = statenormal Or m_iState = statedisabled Then
        tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
      Else
        tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNFACE))
      End If

      UserControl.BackColor = tmpcolor

      If m_iState = statenormal Then
        'Normal (flat)
        tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
        UserControl.BackColor = tmpcolor
        APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
      ElseIf m_iState = stateHot Then
        'Hover
        'tmpColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
        'UserControl.BackColor = tmpColor
        DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_RAISEDINNER
      ElseIf m_iState = statePressed Then
        'Pushed
        'tmpColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
        'UserControl.BackColor = tmpColor
        DrawCtlEdge UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, BDR_SUNKENOUTER
      Else    'Disabled
        'tmpColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
        'UserControl.BackColor = tmpColor
        APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
      End If

      'If Ambient.DisplayAsDefault Then
      '    APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, vbBlack
      'End If
    Case isbJava
      'Java Style
      Dim tmpRect As RECT
      UserControl.BackColor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))

      Select Case m_iState
        Case statePressed
          tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_BTNSHADOW))
        Case stateHot
          tmpcolor = IIf(m_bUseCustomColors, BlendColors(m_lHighlightColor, m_lBackColor), BlendColors(GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_BTNFACE)))
        Case Else
          tmpcolor = IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE))
      End Select
      CopyRect tmpRect, m_btnRect: InflateRect tmpRect, -4, -4
      APIFillRect UserControl.hdc, m_btnRect, tmpcolor 'm_txtRect
      DrawJavaBorder m_btnRect.Left, m_btnRect.Top, m_btnRect.Right - m_btnRect.Left - 1, m_btnRect.Bottom - m_btnRect.Top - 1, GetSysColor(COLOR_BTNSHADOW), GetSysColor(COLOR_WINDOW), tmpcolor
      'If Ambient.DisplayAsDefault Then
      '    APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, vbBlack
      'End If
    Case isbOfficeXP
      'Redmond 2002 Office Suite ( ... )
      UserControl.BackColor = tmpcolor

      If m_iState = statenormal Then
        'If Ambient.DisplayAsDefault Then
        '    tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
        '    APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, MSOXPShiftColor(tmpcolor)
        '    APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
        'End If
        tmpcolor = MSOXPShiftColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H20)
        APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpcolor
      ElseIf m_iState = stateHot Then
        'Hover
        tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
        APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, MSOXPShiftColor(tmpcolor)
        APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
      ElseIf m_iState = statePressed Then
        'Pushed
        tmpcolor = IIf(m_bUseCustomColors, m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
        APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, MSOXPShiftColor(tmpcolor, &H80)
        APIRectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, tmpcolor
      Else
        'Disabled
        tmpcolor = MSOXPShiftColor(IIf(m_bUseCustomColors, m_lBackColor, GetSysColor(COLOR_BTNFACE)), &H20)
        APIFillRectByCoords UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tmpcolor
      End If

    Case isbWindowsXP

      'WinXP (Emulated)
      If m_bUseCustomColors Then
        DrawCustomWinXPButton m_iState
      Else
        DrawWinXPButton m_iState
      End If

      iStyleIconOffset = 3
    Case isbWindowsTheme
      'Uses the current installed windows theme
      Dim bDrawThemeSuccess As Boolean
      Dim tmpStyle As isbStyle
      UserControl.BackColor = GetSysColor(COLOR_BTNFACE)

      If m_iState = (statenormal And m_bFocused) Then 'Or Ambient.DisplayAsDefault Then
        bDrawThemeSuccess = DrawTheme("Button", 1, stateDefaulted)
      Else
        bDrawThemeSuccess = DrawTheme("Button", 1, m_iState)
      End If

      If Not bDrawThemeSuccess Then
        m_iStyle = Me.NonThemeStyle
      End If

    Case isbPlastik
      DrawPlastikButton m_iState
      iStyleIconOffset = 4
    Case isbGalaxy
      DrawGalaxyButton m_iState
      iStyleIconOffset = 4
    Case isbKeramik
      DrawKeramikButton m_iState
      iStyleIconOffset = 6
    Case isbMacOSX
      '�?�??�?Yes! Do you like It?
      DrawMacOSXButton
      iStyleIconOffset = 7
  End Select

    ''''''Draw Icon
    If Not m_Icon Is Nothing Then
        If Icon <> 0 Then
            Dim Ix As Long, Iy As Long
            If m_IconAlign = isbCenter Then
                Ix = (UserControl.ScaleWidth - m_IconSize) / 2
                Iy = (UserControl.ScaleHeight - m_IconSize) / 2
            ElseIf m_IconAlign = isbbottom Then
                Ix = (UserControl.ScaleWidth - m_IconSize) / 2
                Iy = UserControl.ScaleHeight - m_IconSize - iStyleIconOffset
            ElseIf m_IconAlign = isbTop Then
                Ix = (UserControl.ScaleWidth - m_IconSize) / 2
                Iy = iStyleIconOffset
            ElseIf m_IconAlign = isbleft Then
                Ix = iStyleIconOffset
                Iy = (UserControl.ScaleHeight - m_IconSize) / 2
            ElseIf m_IconAlign = isbRight Then
                Ix = UserControl.ScaleWidth - m_IconSize - iStyleIconOffset
                Iy = (UserControl.ScaleHeight - m_IconSize) / 2
            End If
            
            Dim ni As Long, nj As Long
            If m_iState = statePressed Then
                Ix = Ix + 1
                Iy = Iy + 1
            ElseIf m_iState = stateHot Then
                If m_iStyle = isbOfficeXP Then
                    If m_UseMaskColor Then
                        'This was added By t_eee eeee
                        TransBlt UserControl.hdc, Ix + 1, Iy + 1, m_IconSize, m_IconSize, m_Icon, m_MaskColor, &H808080
                        TransBlt UserControl.hdc, Ix - 1, Iy - 1, m_IconSize, m_IconSize, m_Icon, m_MaskColor '                        pMask.PaintPicture m_Icon, ix, iy, m_IconSize, m_IconSize, , , , , vbSrcCopy
                    Else
                        
                        PaintPicture m_Icon, Ix, Iy, m_IconSize, m_IconSize
                        lTransColor = GetPixel(UserControl.hdc, 1, 1)
                        For nj = Iy To Iy + m_IconSize
                            For ni = Ix To Ix + m_IconSize
                                lcurrpix = GetPixel(UserControl.hdc, ni, nj)
                                If lcurrpix <> lTransColor Then
                                    If m_UseMaskColor Then
                                        If lcurrpix <> m_MaskColor Then
                                            SetPixelV UserControl.hdc, ni, nj, &H808080
                                        End If
                                    Else
                                        SetPixelV UserControl.hdc, ni, nj, &H808080
                                    End If
                                End If
                            Next ni
                        Next nj
                    End If
                    Ix = Ix - 2
                    Iy = Iy - 2
                End If
            End If
            'I'll try to mask when usemaskcolor is true
            If m_UseMaskColor Then
                If m_bEnabled Then
                    'Paint in the about pic on color
                    On Error GoTo MalformedIcon
                    TransBlt UserControl.hdc, Ix, Iy, m_IconSize, m_IconSize, m_Icon, m_MaskColor
                Else
                    'Disabled
                    On Error GoTo MalformedIcon
                    TransBlt UserControl.hdc, Ix, Iy, m_IconSize, m_IconSize, m_Icon, m_MaskColor, , , True
                End If
            Else
MalformedIcon:
                If m_bEnabled Then
                   PaintPicture m_Icon, Ix, Iy, m_IconSize, m_IconSize
                Else
                   PaintIconGrayscale UserControl.hdc, m_Icon, Ix, Iy, m_IconSize, m_IconSize
                End If
            End If
        End If
        
    End If
    
    'DrawText
    DrawCaption

End Sub

Private Function lGrayScale(coloredColor As Long) As Long
  'Convertir un Long RGB a un LongRGB grayscale
  'For this task Imported previously developed functions from my project Opticlops
  '' Desc: Convert a RGB color to long
  'Private Function RGBToLong(RGBColor As RGB) As Long
  '    RGBToLong = RGBColor.Blue + RGBColor.Green * 265 + RGBColor.Red * 65536
  'End Function
  '
  '' Desc Convert a long into a RGB structure
  'Private Function LongToRGB(lcolor As Long) As RGB
  '    LongToRGB.Red = lcolor And &HFF
  '    LongToRGB.Green = (lcolor \ &H100) And &HFF
  '    LongToRGB.Blue = (lcolor \ &H10000) And &HFF
  'End Function
  On Error GoTo lGrayScale_Error

  Dim R As Long, G As Long, B As Long
  Dim neutral As Long
  'Splitt into RGB values
  B = coloredColor And &HFF
  G = (coloredColor \ &H100) And &HFF
  R = (coloredColor \ &H10000) And &HFF
  'Obtener el promedio
  neutral = (R / 3 + G / 3 + B / 3)
  'Build Long
  lGrayScale = RGB(neutral, neutral, neutral)
  Exit Function

lGrayScale_Error:
End Function

Private Sub BuildRegion()
  On Error GoTo BuildRegion_Error

  If m_lRegion Then DeleteObject m_lRegion

  Select Case m_iStyle
    Case isbMacOSX
      m_lRegion = CreateMacOSXButtonRegion
    Case isbWindowsXP, isbPlastik
      m_lRegion = CreateWinXPregion
    Case isbGalaxy, isbKeramik
      m_lRegion = CreateGalaxyRegion
    Case Else
      m_lRegion = CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
  End Select

  SetWindowRgn UserControl.hWnd, m_lRegion, True
  Exit Sub

BuildRegion_Error:
End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
  On Error GoTo UserControl_MouseUp_Error
  
  If Me.Enabled = False Then Exit Sub

  If Button = vbLeftButton Then
    If m_ButtonType = isbCheckBox And Not m_Value Then
      m_iState = statePressed
    Else
      m_iState = stateHot
      DoEvents
      Call TrackMouseLeave(UserControl.hWnd)
      If bInCtrl Then RaiseEvent Click
    End If

    Refresh
  End If

  lPrevButton = Button
  Exit Sub

UserControl_MouseUp_Error:
End Sub

Private Sub UserControl_Paint()
  On Error GoTo UserControl_Paint_Error

  Call Refresh
  Exit Sub

UserControl_Paint_Error:
End Sub

Private Sub UserControl_Show()
  On Error GoTo UserControl_Show_Error

  m_bVisible = True
  UserControl_Resize
  Refresh
  Exit Sub

UserControl_Show_Error:
End Sub

Private Sub UserControl_Resize()
  Dim tmpRect As RECT
  Dim lh As Long, lw As Long
  On Error Resume Next
  
  If UserControl.Width < 300 Then UserControl.Width = 300
  If UserControl.Height < 300 Then UserControl.Height = 300

  lh = UserControl.ScaleHeight
  lw = UserControl.ScaleWidth
  SetRect m_btnRect, 0, 0, lw, lh
  SetRect m_txtRect, 4, 4, lw - 4, lh - 4
  CopyRect tmpRect, m_txtRect
  DrawText UserControl.hdc, m_sCaption, Len(m_sCaption), tmpRect, DT_CALCRECT Or DT_WORDBREAK Or IIf(m_bRTLText, DT_RTLREADING, 0)
  
  Select Case m_CaptionAlign
    Case isbCenter
      SetRect m_txtRect, (lw - tmpRect.Right - tmpRect.Left) / 2, (lh - tmpRect.Bottom - tmpRect.Top) / 2 + 3, (lw + tmpRect.Right - tmpRect.Left) / 2, (lh + tmpRect.Bottom - tmpRect.Top) / 2
      lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_WORDBREAK
    Case isbleft
      CopyRect m_txtRect, tmpRect
      SetRect m_txtRect, iStyleIconOffset, (lh - tmpRect.Bottom - tmpRect.Top) / 2 + 3, tmpRect.Right + iStyleIconOffset, (lh + tmpRect.Bottom - tmpRect.Top) / 2
      lwFontAlign = DT_VCENTER Or DT_LEFT Or DT_WORDBREAK
    Case isbRight
      CopyRect m_txtRect, tmpRect
      SetRect m_txtRect, (lw - tmpRect.Right - tmpRect.Left) - iStyleIconOffset, (lh - tmpRect.Bottom - tmpRect.Top) / 2 + 3, (lw - tmpRect.Left) - iStyleIconOffset, (lh + tmpRect.Bottom - tmpRect.Top) / 2
      lwFontAlign = DT_VCENTER Or DT_RIGHT Or DT_WORDBREAK
    Case isbTop
      CopyRect m_txtRect, tmpRect
      SetRect m_txtRect, (lw - tmpRect.Right - tmpRect.Left) / 2, iStyleIconOffset / 2, (lw + tmpRect.Right - tmpRect.Left) / 2, iStyleIconOffset / 2 + (tmpRect.Bottom - tmpRect.Top)
      lwFontAlign = DT_CENTER Or DT_TOP Or DT_WORDBREAK
    Case isbbottom
      CopyRect m_txtRect, tmpRect
      SetRect m_txtRect, (lw - tmpRect.Right - tmpRect.Left) / 2, lh - (tmpRect.Bottom - tmpRect.Top) - iStyleIconOffset / 2, (lw + tmpRect.Right - tmpRect.Left) / 2, lh - iStyleIconOffset / 2
      lwFontAlign = DT_CENTER Or DT_BOTTOM Or DT_WORDBREAK
  End Select
      lwFontAlign = lwFontAlign Or IIf(m_bRTLText, DT_RTLREADING, 0)
  BuildRegion
  Refresh
End Sub

'''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''
'Properties
'Read the properties from the property bag - also, a good place to start the subclassing (if we're running)
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error GoTo UserControl_ReadProperties_Error

  m_iState = statenormal

  With PropBag
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    m_iStyle = PropBag.ReadProperty("Style", 3)
    m_sCaption = PropBag.ReadProperty("Caption", Caption)
    m_IconSize = PropBag.ReadProperty("IconSize", 16)
    m_CaptionAlign = PropBag.ReadProperty("CaptionAlign", isbCenter)
    m_IconAlign = PropBag.ReadProperty("IconAlign", isbleft)
    m_iNonThemeStyle = PropBag.ReadProperty("iNonThemeStyle", [isbWindowsXP])
    m_bEnabled = PropBag.ReadProperty("Enabled", True)
    m_bShowFocus = PropBag.ReadProperty("ShowFocus", False)
    m_bUseCustomColors = PropBag.ReadProperty("USeCustomColors", False)
    m_lBackColor = PropBag.ReadProperty("BackColor", GetSysColor(COLOR_BTNFACE))
    m_lHighlightColor = PropBag.ReadProperty("HighlightColor", GetSysColor(COLOR_HIGHLIGHT))
    m_lFontColor = PropBag.ReadProperty("FontColor", GetSysColor(COLOR_BTNTEXT))
    m_lFontHighlightColor = PropBag.ReadProperty("FontHighlightColor", GetSysColor(COLOR_BTNTEXT))
    UserControl.Extender.ToolTipText = PropBag.ReadProperty("ToolTipText", m_sToolTipText)
    m_sToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_sTooltiptitle = PropBag.ReadProperty("Tooltiptitle", "")
    m_lToolTipIcon = PropBag.ReadProperty("ToolTipIcon", 0)
    m_lToolTipType = PropBag.ReadProperty("ToolTipType", 1)
    m_lttBackColor = PropBag.ReadProperty("ttBackColor", GetSysColor(COLOR_INFOTEXT))
    m_lttForeColor = PropBag.ReadProperty("ttForeColor", GetSysColor(COLOR_INFOBK))
    Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
    m_ButtonType = .ReadProperty("ButtonType", isbButton)
    m_Value = .ReadProperty("Value", False)
    m_UseMaskColor = .ReadProperty("UseMaskColor", False)
    m_MaskColor = .ReadProperty("MaskColor", &HC0C0C0)
    UserControl.MousePointer = .ReadProperty("MousePointer", 0)
    m_bRoundedBordersByTheme = .ReadProperty("RoundedBordersByTheme", True)
    m_bRTLText = .ReadProperty("bRTLText", False)
  End With

  If Ambient.UserMode Then 'If we're not in design mode
    bTrack = True
    bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

    If Not bTrackUser32 Then
      If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
        bTrack = False
      End If
    End If

    If bTrack Then
      'OS supports mouse leave so subclass for it
      With UserControl
        'Start subclassing the UserControl
        Call sc_Subclass(.hWnd)
        Call sc_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_AFTER)
        Call sc_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
        Call sc_AddMsg(.hWnd, WM_THEMECHANGED, MSG_AFTER)
        Call sc_AddMsg(.hWnd, WM_SYSCOLORCHANGE, MSG_AFTER)
      End With

    End If
  End If
  UserControl_Resize
  Exit Sub

UserControl_ReadProperties_Error:
End Sub

'The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
  On Error GoTo Catch

  Call sc_Terminate
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error GoTo UserControl_WriteProperties_Error

  With PropBag
    Call .WriteProperty("Icon", m_Icon)
    Call .WriteProperty("Style", m_iStyle, 3)
    Call .WriteProperty("Caption", m_sCaption, "")
    Call .WriteProperty("IconSize", m_IconSize, 16)
    Call .WriteProperty("IconAlign", m_IconAlign, isbleft)
    Call .WriteProperty("CaptionAlign", m_CaptionAlign, isbCenter)
    Call .WriteProperty("iNonThemeStyle", m_iNonThemeStyle, isbWindowsXP)
    Call .WriteProperty("Enabled", m_bEnabled, True)
    Call .WriteProperty("ShowFocus", m_bShowFocus, False)
    Call .WriteProperty("USeCustomColors", m_bUseCustomColors, False)
    Call .WriteProperty("BackColor", m_lBackColor, GetSysColor(COLOR_BTNFACE))
    Call .WriteProperty("HighlightColor", m_lHighlightColor, GetSysColor(COLOR_HIGHLIGHT))
    Call .WriteProperty("FontColor", m_lFontColor, GetSysColor(COLOR_BTNTEXT))
    Call .WriteProperty("FontHighlightColor", m_lFontHighlightColor, GetSysColor(COLOR_BTNTEXT))
    Call .WriteProperty("ToolTipText", UserControl.Extender.ToolTipText) ' m_sToolTipText, UserControl.Extender.ToolTipText)
    Call .WriteProperty("ToolTipTitle", m_sTooltiptitle)
    Call .WriteProperty("ToolTipIcon", m_lToolTipIcon)
    Call .WriteProperty("ToolTipType", m_lToolTipType)
    Call .WriteProperty("ttBackColor", m_lttBackColor, GetSysColor(COLOR_INFOTEXT))
    Call .WriteProperty("ttForeColor", m_lttForeColor, GetSysColor(COLOR_INFOBK))
    Call .WriteProperty("Font", UserControl.Font)
    Call .WriteProperty("ButtonType", m_ButtonType, isbButton)
    Call .WriteProperty("Value", m_Value, False)
    Call .WriteProperty("MaskColor", m_MaskColor, &HC0C0C0)
    Call .WriteProperty("UseMaskColor", m_UseMaskColor, False)
    Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call .WriteProperty("RoundedBordersByTheme", m_bRoundedBordersByTheme, True)
    Call .WriteProperty("bRTLText", m_bRTLText, False)
  End With

  Exit Sub

UserControl_WriteProperties_Error:
End Sub

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  On Error GoTo TrackMouseLeave_Error

  Dim tme As TRACKMOUSEEVENT_STRUCT

  If bTrack Then

    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If

  Exit Sub

TrackMouseLeave_Error:
End Sub

Public Function Version() As String
  On Error GoTo Version_Error

  Version = strCurrentVersion
  Exit Function

Version_Error:
End Function

' Description: this is the Style property.
Public Property Let Style(ByVal NewStyle As isbStyle)
  On Error GoTo Style_Error

  m_iStyle = NewStyle
  PropertyChanged "Style"
  'For a small error when creating pixarray, I need to set the backcolor to white
  'If m_iStyle = isbMacOSX Then UserControl.BackColor = vbWhite
  UserControl_Resize
  Refresh
  UserControl_Resize
  Refresh
  Exit Property

Style_Error:
End Property

Public Property Get Style() As isbStyle
  On Error GoTo Style_Error

  Style = m_iStyle
  Exit Property

Style_Error:
End Property

' Description: this is the "Caption" property.
Public Property Let Caption(ByVal NewCaption As String)
  On Error GoTo Caption_Error

  m_sCaption = NewCaption
  PropertyChanged "Caption"
  UserControl_Resize
  Refresh
  Exit Property

Caption_Error:
End Property

Public Property Get Caption() As String
  On Error GoTo Caption_Error

  Caption = m_sCaption
  Exit Property

Caption_Error:
End Property

' Description: this is the Picture Property
Public Property Set Icon(NewIcon As StdPicture)
  On Error GoTo Icon_Error

  Set m_Icon = NewIcon
  PropertyChanged "Icon"
  UserControl_Resize
  UserControl_Paint
  Exit Property

Icon_Error:
End Property

Public Property Get Icon() As StdPicture
  On Error GoTo Icon_Error

  Set Icon = m_Icon
  Exit Property

Icon_Error:
End Property

' Description: this is the "IconAlign" property.
Public Property Let IconAlign(ByVal NewIconAlign As isbAlign)
  On Error GoTo IconAlign_Error

  m_IconAlign = NewIconAlign
  PropertyChanged "IconAlign"
  UserControl_Resize
  UserControl_Paint
  Exit Property

IconAlign_Error:
End Property

Public Property Get IconAlign() As isbAlign
  On Error GoTo IconAlign_Error

  IconAlign = m_IconAlign
  Exit Property

IconAlign_Error:
End Property

' Description: this is the "IconSize" property.
Public Property Let IconSize(ByVal NewIconSize As Integer)
  On Error GoTo IconSize_Error

  m_IconSize = NewIconSize
  PropertyChanged "IconSize"
  UserControl_Resize
  UserControl_Paint
  Exit Property

IconSize_Error:
End Property

Public Property Get IconSize() As Integer
  On Error GoTo IconSize_Error

  IconSize = m_IconSize
  Exit Property

IconSize_Error:
End Property

' Description: this is the "CaptionAlign" property.
Public Property Let CaptionAlign(ByVal NewCaptionAlign As isbAlign)
  On Error GoTo CaptionAlign_Error

  m_CaptionAlign = NewCaptionAlign
  PropertyChanged "CaptionAlign"
  UserControl_Resize
  UserControl_Paint
  Exit Property

CaptionAlign_Error:
End Property

Public Property Get CaptionAlign() As isbAlign
  On Error GoTo CaptionAlign_Error

  CaptionAlign = m_CaptionAlign
  Exit Property

CaptionAlign_Error:
End Property

' Description: When Themed Faile, Use this style:
Public Property Let NonThemeStyle(ByVal NewNonThemeStyle As isbStyle)
  On Error GoTo NonThemeStyle_Error

  m_iNonThemeStyle = NewNonThemeStyle
  PropertyChanged "NonThemeStyle"
  UserControl_Resize
  UserControl_Paint
  Exit Property

NonThemeStyle_Error:
End Property

Public Property Get NonThemeStyle() As isbStyle
  On Error GoTo NonThemeStyle_Error

  NonThemeStyle = m_iNonThemeStyle
  Exit Property

NonThemeStyle_Error:
End Property

'Description: Enable or disable the control
Public Property Let Enabled(bEnabled As Boolean)
  On Error GoTo Enabled_Error

  m_bEnabled = bEnabled
  m_iState = statenormal
  Refresh
  PropertyChanged "Enabled"
  UserControl.Enabled = m_bEnabled
  Exit Property

Enabled_Error:
End Property

Public Property Get Enabled() As Boolean
  On Error GoTo Enabled_Error

  Enabled = m_bEnabled
  Refresh
  Exit Property

Enabled_Error:
End Property

'Description: Do we want to show Focus?
Public Property Let ShowFocus(bShowFocus As Boolean)
  On Error GoTo ShowFocus_Error

  m_bShowFocus = bShowFocus
  PropertyChanged "ShowFocus"
  Refresh
  Exit Property

ShowFocus_Error:
End Property

Public Property Get ShowFocus() As Boolean
  On Error GoTo ShowFocus_Error

  ShowFocus = m_bShowFocus
  Exit Property

ShowFocus_Error:
End Property

'Description: Will we use custom colors?
'             If not, system colors will be used
Public Property Let UseCustomColors(bUseCustomColors As Boolean)
  On Error GoTo UseCustomColors_Error

  m_bUseCustomColors = bUseCustomColors
  PropertyChanged "UseCustomColors"
  Refresh
  Exit Property

UseCustomColors_Error:
End Property

Public Property Get UseCustomColors() As Boolean
  On Error GoTo UseCustomColors_Error

  UseCustomColors = m_bUseCustomColors
  Exit Property

UseCustomColors_Error:
End Property

'Description: Use this color for drawing
Public Property Let BackColor(lBackColor As OLE_COLOR)
  On Error GoTo BackColor_Error

  m_lBackColor = lBackColor
  PropertyChanged "BackColor"
  Refresh
  Exit Property

BackColor_Error:
End Property

Public Property Get BackColor() As OLE_COLOR
  On Error GoTo BackColor_Error

  BackColor = m_lBackColor
  Exit Property

BackColor_Error:
End Property

'Description: Use this color for drawing
Public Property Let HighlightColor(lHighlightColor As OLE_COLOR)
  On Error GoTo HighlightColor_Error

  m_lHighlightColor = lHighlightColor
  PropertyChanged "HighlightColor"
  Refresh
  Exit Property

HighlightColor_Error:
End Property

Public Property Get HighlightColor() As OLE_COLOR
  On Error GoTo HighlightColor_Error

  HighlightColor = m_lHighlightColor
  Exit Property

HighlightColor_Error:
End Property

'Description: Use this color for drawing normal font
Public Property Let FontColor(lFontColor As OLE_COLOR)
  On Error GoTo FontColor_Error

  m_lFontColor = lFontColor
  PropertyChanged "FontColor"
  Refresh
  Exit Property

FontColor_Error:
End Property

Public Property Get FontColor() As OLE_COLOR
  On Error GoTo FontColor_Error

  FontColor = m_lFontColor
  Exit Property

FontColor_Error:
End Property

'Description: Use this color for drawing normal font
Public Property Let FontHighlightColor(lFontHighlightColor As OLE_COLOR)
  On Error GoTo FontHighlightColor_Error

  m_lFontHighlightColor = lFontHighlightColor
  PropertyChanged "FontHighlightColor"
  Refresh
  Exit Property

FontHighlightColor_Error:
End Property

Public Property Get FontHighlightColor() As OLE_COLOR
  On Error GoTo FontHighlightColor_Error

  FontHighlightColor = m_lFontHighlightColor
  Exit Property

FontHighlightColor_Error:
End Property

'Description: Set TooltipText
'This sub Is never executed:(
Public Property Let ToolTipText(sToolTipText As String)
  'UserControl.Extender.ToolTipText = sToolTipText
  On Error GoTo ToolTipText_Error

  m_sToolTipText = sToolTipText
  CreateToolTip
  PropertyChanged "ToolTipText"
  Refresh
  Exit Property

ToolTipText_Error:
End Property

Public Property Get ToolTipText() As String
  'UserControl.Extender.ToolTipText = "" 'm_sToolTipText
  'ToolTipText = UserControl.Extender.ToolTipText
  On Error GoTo ToolTipText_Error

  ToolTipText = m_sToolTipText
  Exit Property

ToolTipText_Error:
End Property

'Description: Set TooltipTitle
Public Property Let ToolTipTitle(sTooltipTitle As String)
  On Error GoTo ToolTipTitle_Error

  m_sTooltiptitle = sTooltipTitle
  PropertyChanged "TooltipTitle"
  Refresh
  Exit Property

ToolTipTitle_Error:
End Property

Public Property Get ToolTipTitle() As String
  On Error GoTo ToolTipTitle_Error

  ToolTipTitle = m_sTooltiptitle
  Exit Property

ToolTipTitle_Error:
End Property

'Description: Set TooltipIcon
Public Property Let ToolTipIcon(lTooltipIcon As ttIconType)
  On Error GoTo ToolTipIcon_Error

  m_lToolTipIcon = lTooltipIcon
  PropertyChanged "TooltipIcon"
  Refresh
  Exit Property

ToolTipIcon_Error:
End Property

Public Property Get ToolTipIcon() As ttIconType
  On Error GoTo ToolTipIcon_Error

  ToolTipIcon = m_lToolTipIcon
  Exit Property

ToolTipIcon_Error:
End Property

'Description: Set ToolTipType
Public Property Let ToolTipType(lNewTTType As ttStyleEnum)
  On Error GoTo ToolTipType_Error

  m_lToolTipType = lNewTTType
  PropertyChanged "ToolTipType"
  Refresh
  Exit Property

ToolTipType_Error:
End Property

Public Property Get ToolTipType() As ttStyleEnum
  On Error GoTo ToolTipType_Error

  ToolTipType = m_lToolTipType
  Exit Property

ToolTipType_Error:
End Property

'Description: Set ToolTipBackColor
Public Property Let ToolTipBackColor(lToolTipBackColor As OLE_COLOR)
  On Error GoTo ToolTipBackColor_Error

  m_lttBackColor = lToolTipBackColor
  PropertyChanged "ToolTipBackColor"
  Refresh
  Exit Property

ToolTipBackColor_Error:
End Property

Public Property Get ToolTipBackColor() As OLE_COLOR
  On Error GoTo ToolTipBackColor_Error

  ToolTipBackColor = m_lttBackColor
  Exit Property

ToolTipBackColor_Error:
End Property

'Description: Set ToolTipForeColor
Public Property Let ToolTipForeColor(lToolTipForeColor As OLE_COLOR)
  On Error GoTo ToolTipForeColor_Error

  m_lttForeColor = lToolTipForeColor
  PropertyChanged "ToolTipForeColor"
  Refresh
  Exit Property

ToolTipForeColor_Error:
End Property

Public Property Get ToolTipForeColor() As OLE_COLOR
  On Error GoTo ToolTipForeColor_Error

  ToolTipForeColor = m_lttForeColor
  Exit Property

ToolTipForeColor_Error:
End Property

' Desc: As Requested, Font Property
Public Property Set Font(NewFont As StdFont)
  On Error GoTo Font_Error

  Set m_Font = NewFont
  Set UserControl.Font = NewFont
  Refresh
  PropertyChanged "Font"
  Exit Property

Font_Error:
End Property

Public Property Get Font() As StdFont
  On Error GoTo Font_Error

  Set Font = UserControl.Font
  Exit Property

Font_Error:
End Property

'Description: Set Type Button or CheckBox
Public Property Let ButtonType(newType As isbButtonType)
  On Error GoTo ButtonType_Error

  m_ButtonType = newType
  PropertyChanged "ButtonType"
  Refresh
  Exit Property

ButtonType_Error:
End Property

Public Property Get ButtonType() As isbButtonType
  On Error GoTo ButtonType_Error

  ButtonType = m_ButtonType
  Exit Property

ButtonType_Error:
End Property

'Description: Set the currently used cursor.
Public Property Let MousePointer(lCursor As MousePointerConstants)
  On Error GoTo MousePointer_Error

  UserControl.MousePointer = lCursor
  PropertyChanged "MousePointer"
  Exit Property

MousePointer_Error:
End Property

Public Property Get MousePointer() As MousePointerConstants
  On Error GoTo MousePointer_Error

  MousePointer = UserControl.MousePointer
  Exit Property

MousePointer_Error:
End Property

'Description: Set Button Value (pressed or not pressed)
Public Property Let Value(NewValue As Boolean)
  On Error GoTo Value_Error

  m_Value = NewValue
  PropertyChanged "Value"
  Refresh
  Exit Property

Value_Error:
End Property

Public Property Get Value() As Boolean
  On Error GoTo Value_Error

  Value = m_Value
  Exit Property

Value_Error:
End Property

'Description: Set The Mask Color in the icon
'For BMP Images
Public Property Let MaskColor(NewValue As OLE_COLOR)
  '&H00C0C0C0& Default color
  On Error GoTo MaskColor_Error

  m_MaskColor = NewValue
  PropertyChanged "MaskColor"
  Refresh
  Exit Property

MaskColor_Error:
End Property

Public Property Get MaskColor() As OLE_COLOR
  On Error GoTo MaskColor_Error

  MaskColor = m_MaskColor
  Exit Property

MaskColor_Error:
End Property

Public Property Let UseMaskColor(NewValue As Boolean)
  On Error GoTo UseMaskColor_Error

  m_UseMaskColor = NewValue
  PropertyChanged "UseMaskColor"
  Refresh
  Exit Property

UseMaskColor_Error:
End Property

Public Property Get UseMaskColor() As Boolean
  On Error GoTo UseMaskColor_Error

  UseMaskColor = m_UseMaskColor
  Exit Property

UseMaskColor_Error:
End Property

'Description: Set Rounded Borders If requested By Windows
'Theme
Public Property Let RoundedBordersByTheme(NewValue As Boolean)
    '&H00C0C0C0& Default color
    On Error GoTo RoundedBordersByTheme_Error
    
    m_bRoundedBordersByTheme = NewValue
    PropertyChanged "RoundedBordersByTheme"
    Refresh
    
RoundedBordersByTheme_Error:
End Property

Public Property Get RoundedBordersByTheme() As Boolean
    On Error GoTo RoundedBordersByTheme_Error
    
    RoundedBordersByTheme = m_bRoundedBordersByTheme
    
RoundedBordersByTheme_Error:
End Property

'Description: Use property RightToLeft
'Theme
Public Property Let RightToLeft(NewValue As Boolean)
    '&H00C0C0C0& Default color
    On Error GoTo bRTLText_Error
    
    m_bRTLText = NewValue
    PropertyChanged "bRTLText"
    Refresh
    
bRTLText_Error:
End Property

Public Property Get RightToLeft() As Boolean
    On Error GoTo bRTLText_Error
    
    RightToLeft = m_bRTLText
    
bRTLText_Error:
End Property


'Desc: Opens a url link
'      Call on a Click Event to simulate a hyperlink:
'      isButton1.OpenLink "http://www.geocities.com/fred_cpp/"
'      cmdMail.OpenLink "Mailto:fred_cpp@msn.com"
Public Function OpenLink(sLink As String) As Long
  On Error GoTo OpenLink_Error

  OpenLink = ShellExecute(hWnd, "open", sLink, vbNull, vbNull, 1)
  Exit Function

OpenLink_Error:
End Function
' Fred.cpp  /   2005-July-29   /   4300 lines

'*******************************************************************
' ESTA SUBCLASS, PARA FUNCIONAR, DEVE ESTAR SEMPRE NO FINAL DE TUDO
'*******************************************************************

'*******************************************************************
' THIS SUBCLASS MUST BE AT THE END OF THE CODE TO WORK PROPERLY
'*******************************************************************

'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
  'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data
  'Notes:
  'If you really know what you're doing, it's possible to change the values of the
  'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
  'values get passed to the default handler.. and optionaly, the 'after' callback
  On Error GoTo zSubclass_Proc_Error

  Select Case uMsg
    Case WM_MOUSEMOVE

      If Not bInCtrl Then
        bInCtrl = True
        Call TrackMouseLeave(lng_hWnd)
        m_iState = stateHot
        Refresh
        RaiseEvent MouseEnter
        CreateToolTip
      End If

    Case WM_MOUSELEAVE
      bInCtrl = False
      m_iState = statenormal
      RemoveToolTip
      Refresh
      RaiseEvent MouseLeave
    Case WM_SYSCOLORCHANGE
      Refresh
    Case WM_THEMECHANGED
      Refresh
  End Select

  Exit Sub

zSubclass_Proc_Error:
End Sub
