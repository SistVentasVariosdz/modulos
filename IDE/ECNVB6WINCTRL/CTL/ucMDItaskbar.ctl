VERSION 5.00
Begin VB.UserControl ucMDItaskbar 
   Alignable       =   -1  'True
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   ControlContainer=   -1  'True
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   362
End
Attribute VB_Name = "ucMDItaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------
'Autor:         Leandro Ascierto
'Web:           http://leandroascierto.com/blog/
'Date:          03/08/2011
'Test:          Windows XP, Window Seven
'
'HISTORY:
'           12/09/2011 add WM_PARENTNOTIFY for intercept WM_CREATE in MDICLIENT, this fix when child window is Default Minized or Maximized.
'------------------------------------------------

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const CODE_LEN               As Long = 200                                      'Length of the machine code in bytes
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
  sCode                              As String
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array
Private sc_aBuf(1 To CODE_LEN)       As Byte                                            'Code buffer byte array
Private sc_pCWP                      As Long                                            'Address of the CallWindowsProc
Private sc_pEbMode                   As Long                                            'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                      As Long                                            'Address of the SetWindowsLong function
  
Private Declare Sub RtlMoveMemory Lib "KERNEL32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "KERNEL32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "KERNEL32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
'------------------------------------------------------------------------------------------------


Private Type LOGFONT
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type

Private Const LOGPIXELSY             As Long = 90
Private Const FW_NORMAL              As Long = 400
Private Const FW_BOLD                As Long = 700

Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function MulDiv Lib "KERNEL32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
'== Gdi32
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
'== User32
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function EnableWindow Lib "USER32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetClientRect Lib "USER32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "USER32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "USER32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowPlacement Lib "USER32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "USER32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenuEx Lib "USER32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hwnd As Long, ByVal lptpm As Any) As Long
Private Declare Function GetSystemMenu Lib "USER32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function EndPaint Lib "USER32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function BeginPaint Lib "USER32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function PtInRect Lib "USER32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function DestroyIcon Lib "USER32" (ByVal hIcon As Long) As Long
Private Declare Function RedrawWindow Lib "USER32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DefWindowProc Lib "USER32" Alias "DefWindowProcW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function LoadIcon Lib "USER32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
'== Comctl32
Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hImageList As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_SetImageCount Lib "comctl32.dll" (ByVal himl As Long, ByVal uNewCount As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()


Private Const ILC_MASK          As Long = &H1
Private Const ILC_COLOR32       As Long = &H20
Private Const ILD_TRANSPARENT   As Long = &H1

Private Type TBBUTTON
   iBitmap    As Long
   IdCommand  As Long
   fsState    As Byte
   fsStyle    As Byte
   bReserved1 As Byte
   bReserved2 As Byte
   dwData     As Long
   iString    As Long
End Type

Private Type TBBUTTONINFO
   cbSize    As Long
   dwMask    As Long
   IdCommand As Long
   iImage    As Long
   fsState   As Byte
   fsStyle   As Byte
   cx        As Integer
   lParam    As Long
   pszText   As Long
   cchText   As Long
End Type

Private Type NMHDR
    hwndFrom As Long
    idfrom   As Long
    code     As Long
End Type

Private Type NMTOOLBAR_SHORT
    hdr   As NMHDR
    iItem As Long
End Type

Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type PAINTSTRUCT
    hdc                     As Long
    fErase                  As Long
    rcPaint                 As RECT
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(1 To 32)    As Byte
End Type

Private Type WINDOWPLACEMENT
    Length              As Long
    flags               As Long
    showCmd             As Long
    ptMinPosition       As POINTAPI
    ptMaxPosition       As POINTAPI
    rcNormalPosition    As RECT
End Type

Private Const WPF_SETMINPOSITION    As Long = &H1

Private Const WC_TOOLBAR            As String = "ToolbarWindow32"
Private Const CCS_NODIVIDER         As Long = &H40

Private Const NM_FIRST              As Long = 0
Private Const NM_RCLICK             As Long = (NM_FIRST - 5)
Private Const NM_CLICK              As Long = (NM_FIRST - 2)

Private Const TBIF_IMAGE            As Long = &H1
Private Const TBIF_TEXT             As Long = &H2
Private Const TBIF_STATE            As Long = &H4
Private Const TBIF_LPARAM           As Long = &H10
Private Const TBIF_BYINDEX          As Long = &H80000000

Private Const TBSTYLE_CHECK         As Long = &H2
Private Const TBSTYLE_AUTOSIZE      As Long = &H10
Private Const TBSTYLE_TOOLTIPS      As Long = &H100
Private Const TBSTYLE_WRAPABLE      As Long = &H200
Private Const TBSTYLE_FLAT          As Long = &H800
Private Const TBSTYLE_LIST          As Long = &H1000

Private Const TBSTATE_CHECKED       As Long = &H1
Private Const TBSTATE_PRESSED       As Long = &H2
Private Const TBSTATE_ENABLED       As Long = &H4

Private Const WM_USER               As Long = &H400
Private Const TB_SETSTATE           As Long = (WM_USER + 17)
Private Const TB_GETSTATE           As Long = (WM_USER + 18)
Private Const TB_ADDBITMAP          As Long = (WM_USER + 19)
Private Const TB_ADDBUTTONS         As Long = (WM_USER + 20)
Private Const TB_BUTTONCOUNT        As Long = (WM_USER + 24)
Private Const TB_GETITEMRECT        As Long = (WM_USER + 29)
Private Const TB_BUTTONSTRUCTSIZE   As Long = (WM_USER + 30)
Private Const TB_SETBUTTONSIZE      As Long = (WM_USER + 31)
Private Const TB_SETBITMAPSIZE      As Long = (WM_USER + 32)
Private Const TB_SETIMAGELIST       As Long = (WM_USER + 48)
Private Const TB_GETBUTTONINFOW     As Long = (WM_USER + 63)
Private Const TB_GETBUTTONINFO      As Long = (WM_USER + 65)
Private Const TB_SETBUTTONINFO      As Long = (WM_USER + 66)
Private Const TB_SETEXTENDEDSTYLE   As Long = (WM_USER + 84)
Private Const TB_SETINDENT          As Long = (WM_USER + 47)
Private Const TB_DELETEBUTTON       As Long = (WM_USER + 22)
Private Const TB_INSERTBUTTON       As Long = (WM_USER + 21)
Private Const TB_COMMANDTOINDEX     As Long = (WM_USER + 25)
Private Const TB_GETHOTITEM         As Long = (WM_USER + 71)
Private Const TB_AUTOSIZE           As Long = (WM_USER + 33)
Private Const TB_GETRECT            As Long = (WM_USER + 51)
Private Const TB_GETBUTTONSIZE      As Long = (WM_USER + 58)
Private Const TB_SETBUTTONWIDTH     As Long = (WM_USER + 59)
Private Const TB_SETBUTTONINFOW     As Long = (WM_USER + 64)
Private Const TB_ADDBUTTONSW        As Long = (WM_USER + 68)

Private Const TBN_FIRST             As Long = -700
Private Const TBN_DROPDOWN          As Long = (TBN_FIRST - 10)
Private Const TBN_HOTITEMCHANGE     As Long = (TBN_FIRST - 13)

Private Const TTN_FIRST             As Long = -520
Private Const TTN_GETDISPINFO       As Long = (TTN_FIRST - 0)
Private Const TTM_SETMAXTIPWIDTH    As Long = (WM_USER + 24)

Private Const WM_DESTROY            As Long = &H2
Private Const WM_PAINT              As Long = &HF&
Private Const WM_STYLECHANGED       As Long = &H7D
Private Const WM_SHOWWINDOW         As Long = &H18
Private Const WM_MDINEXT            As Long = &H224
Private Const WM_MDIACTIVATE        As Long = &H222
Private Const WM_GETTEXTLENGTH      As Long = &HE
Private Const WM_GETTEXT            As Long = &HD
Private Const WM_GETICON            As Long = &H7F
Private Const WM_MDIGETACTIVE       As Long = &H229
Private Const WM_SYSCOMMAND         As Long = &H112
Private Const WM_WINDOWPOSCHANGED   As Long = &H47
Private Const WM_SIZE               As Long = &H5
Private Const WM_ERASEBKGND         As Long = &H14
Private Const WM_SETFONT            As Long = &H30
Private Const WM_NOTIFY             As Long = &H4E
Private Const WM_GETFONT            As Long = &H31
Private Const WM_SETICON            As Long = &H80
Private Const WM_SETTEXT            As Long = &HC
Private Const WM_NCUAHDRAWCAPTION   As Long = &HAE
Private Const WM_CREATE             As Long = &H1
Private Const WM_PARENTNOTIFY       As Long = &H210


Private Const SC_RESTORE            As Long = &HF120&
Private Const SC_MAXIMIZE           As Long = &HF030&
Private Const SC_MINIMIZE           As Long = &HF020&
Private Const SC_NEXTWINDOW         As Long = &HF040&

Private Const TPM_RETURNCMD         As Long = &H100&
Private Const TPM_RIGHTALIGN        As Long = &H8&
Private Const TPM_BOTTOMALIGN       As Long = &H20&

Private Const GWL_EXSTYLE           As Long = -20
Private Const WS_EX_MDICHILD        As Long = &H40&
Private Const WS_EX_APPWINDOW       As Long = &H40000

Private Const GWL_STYLE             As Long = (-16)
Private Const WS_VISIBLE            As Long = &H10000000
Private Const WS_CHILD              As Long = &H40000000

Private Const DT_SINGLELINE         As Long = &H20
Private Const DT_VCENTER            As Long = &H4
Private Const DT_WORD_ELLIPSIS      As Long = &H40000

Private Const IDI_APPLICATION       As Long = 32512&
Private Const GA_ROOT               As Long = 2

Public Event Resize()

Private WithEvents m_oFont          As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private hMDIClient                  As Long
Private m_bInitialized              As Boolean
Private m_hToolbar                  As Long
Private m_hImageList                As Long
Private m_hFont                     As Long
Private m_TextNormalColor           As Long
Private m_TextResalteColor          As Long
Private m_TextDisabledColor         As Long
Private m_MaskColor                 As Long
Private hSkin                       As Long
Private m_SkinPicture               As StdPicture
Private m_IconSize                  As Long
Private m_OldSkinBmp                As Long
Private m_ButtonsWidth              As Long
Private m_ButtonsHeight             As Long

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

    On Error Resume Next
  
    Dim uNMHDR  As NMHDR
    Dim uNMTB   As NMTOOLBAR_SHORT
    Dim uRct    As RECT
    Dim hBrush  As Long

    Select Case lng_hWnd

        Case UserControl.hwnd 'UserControl

            Select Case uMsg

                Case WM_NOTIFY
                    
                    Call CopyMemory(uNMHDR, ByVal lParam, Len(uNMHDR))
                    
                    Select Case uNMHDR.code
                     
                        Case NM_CLICK
                            Call CopyMemory(uNMTB, ByVal lParam, Len(uNMTB))
                            pvToolbarButtonClick uNMTB.iItem
                            
                        Case NM_RCLICK
                            Call CopyMemory(uNMTB, ByVal lParam, Len(uNMTB))
                            pvToolbarButtonRightClick uNMTB.iItem

                    End Select

                Case WM_ERASEBKGND

                    If hSkin Then Exit Sub

                    hBrush = CreateSolidBrush(TranslateColor(UserControl.BackColor))
                    
                    If hBrush Then
                        GetClientRect UserControl.hwnd, uRct
                        FillRect wParam, uRct, hBrush
                        DeleteObject hBrush
                    End If
                    
                Case WM_SIZE

                     Call MoveWindow(m_hToolbar, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1)
            
            End Select

        Case hMDIClient 'MDICLIENT


            Select Case uMsg


                Case WM_SIZE
                    pvSetWindowMinimizedPos

                Case WM_PARENTNOTIFY
                    If LoWord(wParam) = WM_CREATE Then
                        Call Subclass_Start(lParam)
                        Call Subclass_AddMsg(lParam, WM_SHOWWINDOW, MSG_AFTER)
                        Call Subclass_AddMsg(lParam, WM_DESTROY, MSG_BEFORE)
                        Call Subclass_AddMsg(lParam, WM_SYSCOMMAND, MSG_BEFORE_AND_AFTER)
                        Call Subclass_AddMsg(lParam, WM_MDIACTIVATE, MSG_AFTER)
                        Call Subclass_AddMsg(lParam, WM_SETICON, MSG_AFTER)
                        Call Subclass_AddMsg(lParam, WM_STYLECHANGED, MSG_AFTER)
                        Call Subclass_AddMsg(lParam, WM_SETTEXT, MSG_AFTER)
                        Call Subclass_AddMsg(lParam, WM_NCUAHDRAWCAPTION, MSG_AFTER)
                    End If
            End Select
            
        Case m_hToolbar 'ToolBar
        
            Select Case uMsg
            
                Case WM_PAINT
                    If hSkin Then pvDrawToolBar
            
                Case WM_ERASEBKGND
                    If hSkin Then bHandled = True
                    
                Case WM_WINDOWPOSCHANGED
                     
                     
                     If Extender.Align <> vbAlignLeft And Extender.Align <> vbAlignRight Then
                        SendMessage m_hToolbar, TB_AUTOSIZE, 0&, ByVal 0&
                        GetWindowRect m_hToolbar, uRct
                        UserControl.Height = (uRct.Bottom - uRct.Top) * Screen.TwipsPerPixelY
                     End If
                     
            End Select
            
        Case Else 'Child Windows
            Dim hRgn As Long
        
            Select Case uMsg
            
                Case WM_DESTROY
                    Subclass_Stop lng_hWnd
                    
                    
                Case WM_SHOWWINDOW
                    
                    If IsIconic(lng_hWnd) Then
                        If GetProp(lng_hWnd, "RgnMin") = 0 Then
                            hRgn = CreateRectRgn(0, 0, 0, 0)
                            SetWindowRgn lng_hWnd, hRgn, True
                            SetProp lng_hWnd, "RgnMin", hRgn
                        End If
                    Else
                        If GetProp(lng_hWnd, "RgnMin") <> 0 Then
                            SetWindowRgn lng_hWnd, 0, False
                            RemoveProp lng_hWnd, "RgnMin"
                        End If
                    End If
                
                    If wParam = 1 Then
                        If Not pvExistWindowButton(lng_hWnd) Then
                            Call pvAddWindow(lng_hWnd)
                        End If
                    Else
                        Call pvRemoveWindow(lng_hWnd)
                        pvSetWindowMinimizedPos
                    End If
                    
                    Call pvFindActive(lng_hWnd)
                    
                Case WM_SYSCOMMAND
                    If bBefore Then
                        pvSetWindowMinimizedPos
                    Else
                        If wParam = SC_MINIMIZE Then
                            If GetProp(lng_hWnd, "RgnMin") = 0 Then
                                 hRgn = CreateRectRgn(0, 0, 0, 0)
                                 SetWindowRgn lng_hWnd, hRgn, True
                                 SetProp lng_hWnd, "RgnMin", hRgn
                            End If
                            pvFindActive SendMessage(hMDIClient, WM_MDIGETACTIVE, 0&, ByVal 0&)
                        Else
                            If GetProp(lng_hWnd, "RgnMin") <> 0 Then
                                SetWindowRgn lng_hWnd, 0, True
                                RemoveProp lng_hWnd, "RgnMin"
                            End If
                        End If
                        
                        If (wParam = SC_NEXTWINDOW) Then
                           
                            SendMessage hMDIClient, WM_MDINEXT, lng_hWnd, ByVal 0&
                            If IsIconic(lng_hWnd) Then
                                hRgn = CreateRectRgn(0, 0, 0, 0)
                                SetWindowRgn lng_hWnd, hRgn, True
                                SetProp lng_hWnd, "RgnMin", hRgn
                            End If
                         
                        End If
                    End If
                    
                Case WM_MDIACTIVATE
                    
                    Call pvFindActive(lParam)
                    pvSetWindowMinimizedPos
                    If IsZoomed(lng_hWnd) Then
                        If GetProp(lng_hWnd, "RgnMin") <> 0 Then
                            SetWindowRgn lng_hWnd, 0, True
                            RemoveProp lng_hWnd, "RgnMin"
                        End If
                    End If

                Case WM_SETICON
                    Call pvChangeIcon(lng_hWnd)
                
                Case WM_STYLECHANGED
                    If Not pvIsShowInTaskBar(lng_hWnd) Then
                        Subclass_Stop lng_hWnd
                    Else
                        Call pvChangeCaption(lng_hWnd)
                    End If
                    
                Case WM_NCUAHDRAWCAPTION, WM_SETTEXT

                    Call pvChangeCaption(lng_hWnd)

            End Select
        
    End Select
End Sub

Private Sub pvToolbarButtonClick(ByVal Button As Long)
    Dim Index As Long
    Dim handle As Long
    Dim hActive As Long
    
    hActive = SendMessage(hMDIClient, WM_MDIGETACTIVE, 0&, ByVal 0&)
    Index = SendMessage(m_hToolbar, TB_COMMANDTOINDEX, Button, ByVal 0&)
    handle = pvButtonParam(Index)
    
    If IsIconic(handle) Then
        If IsZoomed(hActive) Then
            SendMessage handle, WM_SYSCOMMAND, SC_MAXIMIZE, ByVal 0&
        Else
            SendMessage handle, WM_SYSCOMMAND, SC_RESTORE, ByVal 0&
        End If
    Else
        If handle = hActive Then
            SendMessage handle, WM_SYSCOMMAND, SC_MINIMIZE, ByVal 0&
        Else
            BringWindowToTop handle
        End If
    End If

End Sub

Private Sub pvToolbarButtonRightClick(ByVal Button As Long)
    Dim Index As Long
    Dim handle As Long
    Dim lRet As Long
    Dim PT As POINTAPI
    Dim i As Long
    
    Index = SendMessage(m_hToolbar, TB_COMMANDTOINDEX, Button, ByVal 0&)
        
    handle = pvButtonParam(Index)
    BringWindowToTop handle
    GetCursorPos PT
    lRet = TrackPopupMenuEx(GetSystemMenu(handle, False), TPM_RETURNCMD Or TPM_BOTTOMALIGN Or TPM_RIGHTALIGN, PT.X, PT.Y, UserControl.hwnd, ByVal 0&)
    If lRet Then SendMessage handle, WM_SYSCOMMAND, lRet, ByVal 0&
End Sub

Private Sub pvChangeCaption(hwnd As Long)

    Dim i As Long
    For i = 0 To pvButtonCount - 1
        If pvButtonParam(i) = hwnd Then
            pvButtonCaption(i) = pvGetWindowTextW(hwnd)
            Exit For
        End If
    Next
    
End Sub

Private Sub pvChangeIcon(hwnd As Long)
    Dim hIcon As Long
    Dim lCount As Long
    Dim i As Long
    
    hIcon = pvGetWindowIcon(hwnd)
    
    If hIcon = 0 Then
        hIcon = LoadIcon(0&, IDI_APPLICATION)
        lCount = ImageList_AddIcon(m_hImageList, hIcon)
        DestroyIcon hIcon
    Else
        lCount = ImageList_AddIcon(m_hImageList, hIcon)
    End If
    
    For i = 0 To pvButtonCount - 1
        If pvButtonParam(i) = hwnd Then
            pvButtonImage(i) = lCount
            Exit For
        End If
    Next
End Sub

Private Function pvExistWindowButton(hwnd As Long) As Boolean
    Dim i As Long

    For i = 0 To pvButtonCount - 1
        If pvButtonParam(i) = hwnd Then
            pvExistWindowButton = True
            Exit Function
        End If
    Next

End Function

Private Sub pvAddWindow(hwnd As Long)
    Dim sCaption As String
    Dim hIcon As Long
    Dim lCount As Long
    
    sCaption = pvGetWindowTextW(hwnd)

    hIcon = pvGetWindowIcon(hwnd)
    
    If hIcon = 0 Then
        hIcon = LoadIcon(0&, IDI_APPLICATION)
        lCount = ImageList_AddIcon(m_hImageList, hIcon)
        DestroyIcon hIcon
    Else
        lCount = ImageList_AddIcon(m_hImageList, hIcon)
    End If

    pvAddButton sCaption, lCount, hwnd
End Sub

Private Sub pvRemoveWindow(hwnd As Long)
    Dim i As Long
    Dim lCount As Long
    
    lCount = pvButtonCount
    
    For i = 0 To lCount - 1
        If pvButtonParam(i) = hwnd Then
            pvDeleteButton i
            Exit For
        End If
    Next

    If lCount = 1 Then Call ImageList_SetImageCount(m_hImageList, 0&)
End Sub


Private Function pvFindActive(hActive As Long)

    Dim i As Long

    For i = 0 To pvButtonCount - 1
        If pvButtonParam(i) = hActive Then
            If IsIconic(hActive) = 0 Then
                pvButtonCheked(i) = True
            Else
                pvButtonCheked(i) = False
            End If
        Else
            pvButtonCheked(i) = False
        End If
    Next
    
End Function

Private Function pvSetWindowMinimizedPos()

    Dim REC1 As RECT
    Dim REC2 As RECT
    Dim REC3 As RECT
    Dim REC4 As RECT
    Dim handle As Long
    Dim i As Long
    Dim WP As WINDOWPLACEMENT


    GetWindowRect m_hToolbar, REC1
    GetWindowRect hMDIClient, REC2
    
    For i = 0 To pvButtonCount - 1
        handle = pvButtonParam(i)
        SendMessage m_hToolbar, TB_GETITEMRECT, i, REC3
        GetWindowPlacement handle, WP
        WP.flags = WPF_SETMINPOSITION
        
        Select Case UserControl.Extender.Align
            Case vbAlignTop
                WP.ptMinPosition.X = REC1.Left - REC2.Left + REC3.Left
                WP.ptMinPosition.Y = 0

            Case vbAlignBottom
                WP.ptMinPosition.X = REC1.Left - REC2.Left + REC3.Left
                If IsIconic(handle) Then
                    GetWindowRect handle, REC4
                    WP.ptMinPosition.Y = (REC2.Bottom - REC2.Top) - (REC4.Bottom - REC4.Top) - 10
                Else
                    WP.ptMinPosition.Y = (REC2.Bottom - REC2.Top) - 30
                End If
            Case vbAlignRight
                WP.ptMinPosition.Y = REC1.Top - REC2.Top + REC3.Top
                If IsIconic(handle) Then
                    GetWindowRect handle, REC4
                    WP.ptMinPosition.X = (REC2.Right - REC2.Left) - (REC4.Right - REC4.Left) - 10
                Else
                    WP.ptMinPosition.X = (REC2.Right - REC2.Left) - 300
                End If
            Case vbAlignLeft
                WP.ptMinPosition.X = 0
                WP.ptMinPosition.Y = REC1.Top - REC2.Top + REC3.Top
            Case Else
                If REC1.Top > REC2.Top Then
                    WP.ptMinPosition.X = REC1.Left - REC2.Left + REC3.Left
                    If IsIconic(handle) Then
                        GetWindowRect handle, REC4
                        WP.ptMinPosition.Y = (REC2.Bottom - REC2.Top) - (REC4.Bottom - REC4.Top) - 10
                    Else
                        WP.ptMinPosition.Y = (REC2.Bottom - REC2.Top) - 30
                    End If
                Else
                    WP.ptMinPosition.X = REC1.Left - REC2.Left + REC3.Left
                    WP.ptMinPosition.Y = 0
                End If
                
        End Select
        

        
        
        SetWindowPlacement handle, WP
    Next
                          
End Function

'========================================================================================
' Usercontrol
'========================================================================================

Private Sub UserControl_Initialize()
    Call InitCommonControls
    Set m_oFont = New StdFont
    m_IconSize = 32
End Sub


Private Sub UserControl_InitProperties()
    On Error Resume Next
    m_IconSize = 32
    Extender.Align = vbAlignBottom
    Set m_oFont = Ambient.Font
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        Call Me.Initialize(GetAncestor(UserControl.hwnd, GA_ROOT))
    End If
End Sub

Private Sub UserControl_Terminate()
  
  On Error GoTo errH
  
    Call pvDestroyFont
    
    If (m_bInitialized) Then
        '-- Stop subclassing and destroy all
        Call Subclass_StopAll
        Call pvDestroyImageList
        Call pvDestroyToolbar
    End If
errH:
End Sub

Private Sub pvAddButton(Caption As String, Image As Long, ItemData As Long)
    Dim uTBB   As TBBUTTON
    Dim lCount As Long

    lCount = SendMessage(m_hToolbar, TB_BUTTONCOUNT, 0&, ByVal 0&)

    With uTBB
        .IdCommand = lCount
        .iString = StrPtr(Caption)
        .iBitmap = Image
        .fsStyle = TBSTYLE_CHECK
        .fsState = TBSTATE_ENABLED
        .dwData = ItemData
    End With
    
    SendMessage m_hToolbar, TB_ADDBUTTONSW, 1, uTBB
    SendMessage m_hToolbar, TB_SETBUTTONSIZE, 0&, ByVal MakeDWord(m_ButtonsWidth, m_ButtonsHeight)
    SendMessage m_hToolbar, TB_AUTOSIZE, 0&, ByVal 0&
    
    
End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hpal As Long = 0) As Long
    If OleTranslateColor(oClr, hpal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

Private Property Get pvButtonCount() As Long
        pvButtonCount = SendMessage(m_hToolbar, TB_BUTTONCOUNT, 0&, ByVal 0&)
End Property

Private Property Let pvButtonCheked(ByVal Button As Long, ByVal Value As Boolean)
  
  Dim uTBBI As TBBUTTONINFO
    
    If (m_hToolbar) Then
    
        With uTBBI
            .cbSize = Len(uTBBI)
            .dwMask = TBIF_STATE Or TBIF_BYINDEX
        End With
        
        Call SendMessage(m_hToolbar, TB_GETBUTTONINFO, Button, uTBBI)
    
        If Value Then
            uTBBI.fsState = uTBBI.fsState Or TBSTATE_CHECKED
        Else
            uTBBI.fsState = uTBBI.fsState And Not TBSTATE_CHECKED
        End If
        
        Call SendMessage(m_hToolbar, TB_SETBUTTONINFO, Button, uTBBI)
        
    
    End If
End Property


Private Property Let pvButtonCaption(ByVal Button As Long, ByVal Caption As String)

    Dim uTBBI As TBBUTTONINFO

    With uTBBI
        .cbSize = Len(uTBBI)
        .dwMask = TBIF_TEXT Or TBIF_BYINDEX
        .pszText = StrPtr(Caption)
    End With

    Call SendMessage(m_hToolbar, TB_SETBUTTONINFOW, Button, uTBBI)

End Property

Private Function pvDeleteButton(ByVal Button As Long) As Boolean
    pvDeleteButton = SendMessage(m_hToolbar, TB_DELETEBUTTON, Button, ByVal 0&)
    SendMessage m_hToolbar, TB_AUTOSIZE, 0&, ByVal 0&
End Function

Private Property Let pvButtonParam(ByVal Button As Long, ByVal ItemData As Long)
    Dim uTBBI As TBBUTTONINFO

    With uTBBI
        .cbSize = Len(uTBBI)
        .dwMask = TBIF_LPARAM Or TBIF_BYINDEX
        .lParam = ItemData
    End With
    Call SendMessage(m_hToolbar, TB_SETBUTTONINFO, Button, uTBBI)

End Property

Private Property Get pvButtonParam(ByVal Button As Long) As Long
    Dim uTBBI As TBBUTTONINFO

    With uTBBI
        .cbSize = Len(uTBBI)
        .dwMask = TBIF_LPARAM Or TBIF_BYINDEX
    End With
    Call SendMessage(m_hToolbar, TB_GETBUTTONINFO, Button, uTBBI)
    
    pvButtonParam = uTBBI.lParam

End Property

Private Property Let pvButtonImage(ByVal Button As Long, ByVal Image As Long)
    Dim uTBBI As TBBUTTONINFO

    With uTBBI
        .cbSize = Len(uTBBI)
        .dwMask = TBIF_IMAGE Or TBIF_BYINDEX
        .iImage = Image
    End With
    Call SendMessage(m_hToolbar, TB_SETBUTTONINFO, Button, uTBBI)

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_IconSize = .ReadProperty("IconSize", 32)
        m_ButtonsWidth = .ReadProperty("ButtonsWidth", 0)
        m_ButtonsHeight = .ReadProperty("ButtonsHeight", 0)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set Me.SkinPicture = .ReadProperty("SkinPicture", Nothing)
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
        UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace)   'vbButtonFace
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("IconSize", m_IconSize, 32)
        Call .WriteProperty("ButtonsWidth", m_ButtonsWidth, 0)
        Call .WriteProperty("ButtonsHeight", m_ButtonsHeight, 0)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("SkinPicture", m_SkinPicture, Nothing)
        Call .WriteProperty("Font", m_oFont, Ambient.Font)
        Call .WriteProperty("BackColor", UserControl.BackColor, vbButtonFace)   'vbButtonFace
    End With
End Sub


Public Function Initialize(ByVal hwndMDI As Long) As Boolean

        If (m_bInitialized = False) Then

            hMDIClient = FindWindowEx(hwndMDI, ByVal 0&, "MDIClient", vbNullString)
  
            If hMDIClient Then
                If m_IconSize = 0 Then m_IconSize = 32

                If pvCreateToolbar32(m_IconSize) Then

                    m_bInitialized = True

                    '-- Subclass
                    Call Subclass_Start(UserControl.hwnd)
                    Call Subclass_AddMsg(UserControl.hwnd, WM_SIZE, MSG_BEFORE)
                    Call Subclass_AddMsg(UserControl.hwnd, WM_NOTIFY, MSG_BEFORE)
                    Call Subclass_AddMsg(UserControl.hwnd, WM_ERASEBKGND, MSG_BEFORE)
                    
                    Call Subclass_Start(hMDIClient)
                    Call Subclass_AddMsg(hMDIClient, WM_SIZE, MSG_AFTER)
                    Call Subclass_AddMsg(hMDIClient, WM_PARENTNOTIFY, MSG_AFTER)
           
                  
                    
                
                    
'Private Const WM_MDITILE As Long = &H226
'Private Const WM_MDISETMENU As Long = &H230
'Private Const WM_MDIRESTORE As Long = &H223
'Private Const WM_MDIREFRESHMENU As Long = &H234
'Private Const WM_MDINEXT As Long = &H224
'Private Const WM_MDIMAXIMIZE As Long = &H225
'Private Const WM_MDIICONARRANGE As Long = &H228

                
                    Call Subclass_Start(m_hToolbar)
                    Call Subclass_AddMsg(m_hToolbar, WM_ERASEBKGND, MSG_BEFORE)
                    Call Subclass_AddMsg(m_hToolbar, WM_PAINT, MSG_BEFORE)
                    Call Subclass_AddMsg(m_hToolbar, WM_WINDOWPOSCHANGED, MSG_AFTER)
    
                    SendMessage m_hToolbar, TB_AUTOSIZE, 0&, ByVal 0&
    
                    Initialize = True

                End If
            End If
        End If
End Function

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
    
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    If m_hToolbar Then RedrawWindow m_hToolbar, ByVal 0&, ByVal 0&, &H1
End Property

Public Property Get ButtonsWidth() As Long
    ButtonsWidth = m_ButtonsWidth
End Property

Public Property Let ButtonsWidth(ByVal New_Value As Long)
    m_ButtonsWidth = New_Value
    PropertyChanged "ButtonsWidth"
    If m_hToolbar Then
        SendMessage m_hToolbar, TB_SETBUTTONSIZE, 0&, ByVal MakeDWord(m_ButtonsWidth, m_ButtonsHeight)
        SendMessage m_hToolbar, TB_AUTOSIZE, 0&, ByVal 0&
    End If
End Property

Public Property Get ButtonsHeight() As Long
    ButtonsHeight = m_ButtonsHeight
End Property

Public Property Let ButtonsHeight(ByVal New_Value As Long)
    m_ButtonsHeight = New_Value
    PropertyChanged "ButtonsHeight"
    If m_hToolbar Then
        SendMessage m_hToolbar, TB_SETBUTTONSIZE, 0&, ByVal MakeDWord(m_ButtonsWidth, m_ButtonsHeight)
        SendMessage m_hToolbar, TB_AUTOSIZE, 0&, ByVal 0&
    End If
End Property

Public Property Get hwndToolbar() As Long
    hwndToolbar = m_hToolbar
End Property

Public Property Get hwndUserControl() As Long
    hwndUserControl = UserControl.hwnd
End Property

Public Property Get hwndMDIClient() As Long
    hwndMDIClient = hMDIClient
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    If m_hToolbar Then Call EnableWindow(m_hToolbar, New_Enabled)
End Property

Public Property Get IconSize() As Long
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_Size As Long)
    Dim i As Long
    Dim hIcon As Long
    Dim lCount As Long
    
    m_IconSize = New_Size
    
    If m_hToolbar Then
        If m_hImageList Then ImageList_Destroy m_hImageList
    
        m_hImageList = ImageList_Create(m_IconSize, m_IconSize, ILC_MASK Or ILC_COLOR32, 0, 0)
        
        If m_hImageList Then
            Call SendMessage(m_hToolbar, TB_SETIMAGELIST, 0&, ByVal m_hImageList)
        
            For i = 0 To pvButtonCount - 1
        
                hIcon = pvGetWindowIcon(pvButtonParam(i))
                
                If hIcon = 0 Then
                    hIcon = LoadIcon(0&, IDI_APPLICATION)
                    lCount = ImageList_AddIcon(m_hImageList, hIcon)
                    DestroyIcon hIcon
                Else
                    lCount = ImageList_AddIcon(m_hImageList, hIcon)
                End If
    
                pvButtonImage(i) = lCount
            
            Next
        End If
    End If
 
End Property

Public Property Get Font() As StdFont
    Set Font = m_oFont
End Property

Public Property Set Font(ByVal New_Font As StdFont)

    Dim uLF   As LOGFONT
    Dim lChar As Long
    
    Set m_oFont = New_Font
    PropertyChanged "Font"

    With m_oFont
        For lChar = 1 To Len(.Name)
            uLF.lfFaceName(lChar - 1) = CByte(Asc(Mid$(.Name, lChar, 1)))
        Next lChar
        uLF.lfHeight = -MulDiv(.SIZE, GetDeviceCaps(UserControl.hdc, LOGPIXELSY), 72)
        uLF.lfItalic = .Italic
        uLF.lfWeight = IIf(.Bold, FW_BOLD, FW_NORMAL)
        uLF.lfUnderline = .Underline
        uLF.lfStrikeOut = .Strikethrough
        uLF.lfCharSet = .Charset
    End With
    
    Call pvDestroyFont: m_hFont = CreateFontIndirect(uLF)
        
    If m_hToolbar Then Call SendMessage(m_hToolbar, WM_SETFONT, m_hFont, ByVal 1&)

End Property

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Set Font = m_oFont
End Sub

Public Property Set SkinPicture(ByVal NewPic As StdPicture)

    If hSkin Then
        Call SelectObject(hSkin, m_OldSkinBmp)
        Call DeleteDC(hSkin): hSkin = 0
    End If
    
    Set m_SkinPicture = NewPic
    
    If Not m_SkinPicture Is Nothing Then
        hSkin = CreateCompatibleDC(0)
        m_OldSkinBmp = SelectObject(hSkin, m_SkinPicture.handle)
        m_TextNormalColor = GetPixel(hSkin, 90, 0)
        m_TextResalteColor = GetPixel(hSkin, 90, 6)
        m_TextDisabledColor = GetPixel(hSkin, 90, 10)
        m_MaskColor = GetPixel(hSkin, 90, 12)
        If m_MaskColor = 0 Then m_MaskColor = -1 'Not use Black Color to the mask
    Else
        Me.ButtonsWidth = m_ButtonsWidth + 1
        Me.ButtonsWidth = m_ButtonsWidth - 1
    End If

    
    If m_hToolbar Then RedrawWindow m_hToolbar, ByVal 0&, ByVal 0&, &H1
    PropertyChanged "SkinPicture"
End Property

Public Property Get SkinPicture() As StdPicture
    Set SkinPicture = m_SkinPicture
End Property

Public Sub SetIndent(ByVal Value As Long)
    If m_hToolbar Then
        Call SendMessage(m_hToolbar, TB_SETINDENT, Value, ByVal 0&)
    End If
End Sub


Private Function pvCreateToolbar32(ByVal ImageSize As Long) As Boolean

    Dim lStyle  As Long
    Dim uButton As TBBUTTON
    Dim lRet As Long
    

    lStyle = WS_CHILD Or WS_VISIBLE Or CCS_NODIVIDER Or TBSTYLE_FLAT Or TBSTYLE_LIST Or TBSTYLE_TOOLTIPS Or TBSTYLE_WRAPABLE
    
    m_hToolbar = CreateWindowEx(0, WC_TOOLBAR, vbNullString, lStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hwnd, 0, App.hInstance, ByVal 0)
    
    If m_hToolbar Then
        
        m_hImageList = ImageList_Create(ImageSize, ImageSize, ILC_MASK Or ILC_COLOR32, 0, 0)
        
        If m_hImageList Then Call SendMessage(m_hToolbar, TB_SETIMAGELIST, 0&, ByVal m_hImageList)
    
        SendMessage m_hToolbar, TB_BUTTONSTRUCTSIZE, Len(uButton), ByVal 0&
   
        SendMessage m_hToolbar, WM_SETFONT, m_hFont, ByVal 0&

        lRet = SendMessage(m_hToolbar, TB_GETBUTTONSIZE, 0&, 0&)
        
        If m_ButtonsWidth = 0 Then m_ButtonsWidth = LoWord(lRet)
        If m_ButtonsHeight = 0 Then m_ButtonsHeight = HiWord(lRet)
        
        
        SendMessage m_hToolbar, TB_SETBUTTONWIDTH, 0&, ByVal MakeDWord(16, 16) 'Force show the ToolTips

        SendMessage m_hToolbar, TB_SETBUTTONSIZE, 0&, ByVal MakeDWord(m_ButtonsWidth, m_ButtonsHeight)
              
        pvCreateToolbar32 = True
    End If
End Function

Private Sub pvDestroyToolbar()
    If (m_hToolbar) Then
        If (DestroyWindow(m_hToolbar)) Then
            m_hToolbar = 0
        End If
    End If
End Sub

Private Sub pvDestroyImageList()
    If (m_hImageList) Then
        If (ImageList_Destroy(m_hImageList)) Then
            m_hImageList = 0
        End If
    End If
End Sub

Private Sub pvDestroyFont()
    If (m_hFont) Then
        If (DeleteObject(m_hFont)) Then
            m_hFont = 0
        End If
    End If
End Sub

Private Function pvGetWindowIcon(hwnd As Long) As Long
    If m_IconSize = 32 Then
        pvGetWindowIcon = SendMessage(hwnd, WM_GETICON, 1&, ByVal 0&)
        If pvGetWindowIcon = 0 Then
            pvGetWindowIcon = SendMessage(hwnd, WM_GETICON, 0&, ByVal 0&)
        End If
    Else
        pvGetWindowIcon = SendMessage(hwnd, WM_GETICON, 0&, ByVal 0&)
        If pvGetWindowIcon = 0 Then
            pvGetWindowIcon = SendMessage(hwnd, WM_GETICON, 1&, ByVal 0&)
        End If
    End If
End Function

Private Function pvGetWindowTextW(hwnd As Long) As String
    Dim strLen As Long
    strLen = DefWindowProc(hwnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    pvGetWindowTextW = String(strLen, 0)
    DefWindowProc hwnd, WM_GETTEXT, Len(pvGetWindowTextW) + 1, ByVal StrPtr(pvGetWindowTextW)
End Function

Private Function pvIsShowInTaskBar(hwnd As Long) As Boolean
    Dim WinExStyle As Long
    WinExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    pvIsShowInTaskBar = WinExStyle = (WinExStyle Or WS_EX_APPWINDOW)
End Function

Private Function LoWord(ByVal Numero As Long) As Long
    LoWord = Numero And &HFFFF&
End Function
 
Private Function HiWord(ByVal Numero As Long) As Long
    HiWord = Numero \ &H10000 And &HFFFF&
End Function
 
Private Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

Private Sub pvDrawToolBar()
    Dim DC As Long
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim PS As PAINTSTRUCT
    Dim TBI As TBBUTTONINFO
    Dim i As Long
    Dim Rec As RECT
    Dim RecText As RECT
    Dim sBuff As String * 260
    Dim HotItem As Long
    Dim PosPressLeftTop As Long
    Dim OldhBmp As Long
    Dim OldhFont As Long

    Call BeginPaint(m_hToolbar, PS)
    
    HotItem = SendMessage(m_hToolbar, TB_GETHOTITEM, 0&, ByVal 0&)
    
    
    DC = GetDC(0)
    hDCMemory = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, UserControl.ScaleWidth, UserControl.ScaleHeight)
    OldhBmp = SelectObject(hDCMemory, hBmp)
    ReleaseDC 0&, DC
    SetStretchBltMode hDCMemory, 4
    SetBkMode hDCMemory, 1 'TRANSPARENT
    
    OldhFont = SelectObject(hDCMemory, m_hFont)
    
    RenderStretchFromDC hDCMemory, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, hSkin, 75, 0, 15, 23, 4, m_MaskColor
    
    With TBI
        .cbSize = Len(TBI)
        .dwMask = TBIF_IMAGE Or TBIF_STATE Or TBIF_TEXT Or TBIF_BYINDEX
        .cchText = 260
        .pszText = StrPtr(sBuff)
    End With
    
    For i = 0 To pvButtonCount - 1
    

        Call SendMessage(m_hToolbar, TB_GETBUTTONINFOW, i, TBI)
        SendMessage m_hToolbar, TB_GETITEMRECT, i, Rec
        PosPressLeftTop = 0
    
        If (TBI.fsState And TBSTATE_CHECKED) = TBSTATE_CHECKED Then
            If (TBI.fsState And TBSTATE_PRESSED) = TBSTATE_PRESSED Then
                RenderStretchFromDC hDCMemory, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, hSkin, 45, 0, 15, 23, 4, m_MaskColor
                PosPressLeftTop = 1
            Else
                If HotItem = i Then
                    RenderStretchFromDC hDCMemory, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, hSkin, 60, 0, 15, 23, 4, m_MaskColor
                Else
                    RenderStretchFromDC hDCMemory, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, hSkin, 45, 0, 15, 23, 4, m_MaskColor
                End If
            End If
            SetTextColor hDCMemory, m_TextResalteColor
        Else
            If (TBI.fsState And TBSTATE_PRESSED) = TBSTATE_PRESSED Then
                SetTextColor hDCMemory, m_TextResalteColor
                PosPressLeftTop = 1
                RenderStretchFromDC hDCMemory, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, hSkin, 30, 0, 15, 23, 4, m_MaskColor
            Else
                SetTextColor hDCMemory, m_TextNormalColor
                If HotItem = i Then
                    RenderStretchFromDC hDCMemory, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, hSkin, 15, 0, 15, 23, 4, m_MaskColor
                Else
                    RenderStretchFromDC hDCMemory, Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top, hSkin, 0, 0, 15, 23, 4, m_MaskColor
                End If
            End If
        End If
    
        With RecText
            .Left = Rec.Left + m_IconSize + PosPressLeftTop + 6
            .Top = Rec.Top + 2 + PosPressLeftTop
            .Right = Rec.Right + PosPressLeftTop - 2
            .Bottom = Rec.Bottom - 2
        End With
        
        DrawTextW hDCMemory, TBI.pszText, -1, RecText, DT_SINGLELINE Or DT_VCENTER Or DT_WORD_ELLIPSIS
        ImageList_Draw m_hImageList, TBI.iImage, hDCMemory, Rec.Left + 4 + PosPressLeftTop, Rec.Top + ((Rec.Bottom - Rec.Top) / 2) - (m_IconSize / 2) + PosPressLeftTop, ILD_TRANSPARENT
      
    Next

    BitBlt PS.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, hDCMemory, 0, 0, vbSrcCopy
    
    Call EndPaint(m_hToolbar, PS)

    Call SelectObject(hDCMemory, OldhFont)
    DeleteObject SelectObject(hDCMemory, OldhBmp)
    DeleteDC hDCMemory

End Sub

Private Function RenderStretchFromDC(ByVal DestDC As Long, _
                                ByVal DestX As Long, _
                                ByVal DestY As Long, _
                                ByVal DestW As Long, _
                                ByVal DestH As Long, _
                                ByVal SrcDC As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                ByVal SIZE As Long, _
                                Optional MaskColor As Long = -1)
 
    Dim Sx2 As Long
     
    Sx2 = SIZE * 2
     
    If MaskColor <> -1 Then
        Dim mDC         As Long
        Dim mX          As Long
        Dim mY          As Long
        Dim DC          As Long
        Dim hBmp        As Long
        Dim hOldBmp     As Long
     
        mDC = DestDC
        DC = GetDC(0)
        DestDC = CreateCompatibleDC(0)
        hBmp = CreateCompatibleBitmap(DC, DestW, DestH)
        hOldBmp = SelectObject(DestDC, hBmp) ' save the original BMP for later reselection
        mX = DestX: mY = DestY
        DestX = 0: DestY = 0
    End If
     
    'SetStretchBltMode DestDC, vbPaletteModeNone
     
    BitBlt DestDC, DestX, DestY, SIZE, SIZE, SrcDC, X, Y, vbSrcCopy  'TOP_LEFT
    StretchBlt DestDC, DestX + SIZE, DestY, DestW - Sx2, SIZE, SrcDC, X + SIZE, Y, Width - Sx2, SIZE, vbSrcCopy 'TOP_CENTER
    BitBlt DestDC, DestX + DestW - SIZE, DestY, SIZE, SIZE, SrcDC, X + Width - SIZE, Y, vbSrcCopy 'TOP_RIGHT
    StretchBlt DestDC, DestX, DestY + SIZE, SIZE, DestH - Sx2, SrcDC, X, Y + SIZE, SIZE, Height - Sx2, vbSrcCopy 'MID_LEFT
    StretchBlt DestDC, DestX + SIZE, DestY + SIZE, DestW - Sx2, DestH - Sx2, SrcDC, X + SIZE, Y + SIZE, Width - Sx2, Height - Sx2, vbSrcCopy 'MID_CENTER
    StretchBlt DestDC, DestX + DestW - SIZE, DestY + SIZE, SIZE, DestH - Sx2, SrcDC, X + Width - SIZE, Y + SIZE, SIZE, Height - Sx2, vbSrcCopy 'MID_RIGHT
    BitBlt DestDC, DestX, DestY + DestH - SIZE, SIZE, SIZE, SrcDC, X, Y + Height - SIZE, vbSrcCopy 'BOTTOM_LEFT
    StretchBlt DestDC, DestX + SIZE, DestY + DestH - SIZE, DestW - Sx2, SIZE, SrcDC, X + SIZE, Y + Height - SIZE, Width - Sx2, SIZE, vbSrcCopy   'BOTTOM_CENTER
    BitBlt DestDC, DestX + DestW - SIZE, DestY + DestH - SIZE, SIZE, SIZE, SrcDC, X + Width - SIZE, Y + Height - SIZE, vbSrcCopy 'BOTTOM_RIGHT
    
    If MaskColor <> -1 Then
        GdiTransparentBlt mDC, mX, mY, DestW, DestH, DestDC, 0, 0, DestW, DestH, MaskColor
        SelectObject DestDC, hOldBmp
        DeleteObject hBmp
        DeleteDC DC
        DeleteDC DestDC
    End If
 
End Function

'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_??? routines
'========================================================================================

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lhWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lhWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
'Private Sub Subclass_DelMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
''Parameters:
'  'lhWnd  - The handle of the window for which the uMsg is to be removed from the callback table
'  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
'  'When      - Whether the msg is to be removed from the before, after or both callback tables
'  With sc_aSubData(zIdx(lhWnd))
'    If When And eMsgWhen.MSG_BEFORE Then
'      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
'    End If
'    If When And eMsgWhen.MSG_AFTER Then
'      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
'    End If
'  End With
'End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lhWnd As Long) As Long
'Parameters:
  'lhWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Dim i                       As Long                                                   'Loop index
  Dim J                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sSubCode                As String                                                 'Subclass code string
Const PUB_CLASSES             As Long = 0                                               'The number of UserControl public classes
Const GMEM_FIXED              As Long = 0                                               'Fixed memory GlobalAlloc flag
Const PAGE_EXECUTE_READWRITE  As Long = &H40&                                           'Allow memory to execute without violating XP SP2 Data Execution Prevention
Const PATCH_01                As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
Const PATCH_02                As Long = 68                                              'Address of the previous WndProc
Const PATCH_03                As Long = 78                                              'Relative address of SetWindowsLong
Const PATCH_06                As Long = 116                                             'Address of the previous WndProc
Const PATCH_07                As Long = 121                                             'Relative address of CallWindowProc
Const PATCH_0A                As Long = 186                                             'Address of the owner object
Const FUNC_CWP                As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
Const FUNC_EBM                As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_SWL                As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const MOD_USER                As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
Const MOD_VBA5                As String = "vba5"                                        'Location of the EbMode function if running VB5
Const MOD_VBA6                As String = "vba6"                                        'Location of the EbMode function if running VB6

'If it's the first time through here..
  If sc_aBuf(1) = 0 Then

'Build the hex pair subclass string
    sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90" & _
               Hex$(&HA4 + (PUB_CLASSES * 12)) & "070000C3"
    
'Convert the string from hex pairs to bytes and store in the machine code buffer
    i = 1
    Do While J < CODE_LEN
      J = J + 1
      sc_aBuf(J) = CByte("&H" & Mid$(sSubCode, i, 2))                                   'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      sc_aBuf(16) = &H90                                                                'Patch the code buffer to enable the IDE state code
      sc_aBuf(17) = &H90                                                                'Patch the code buffer to enable the IDE state code
      sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                        'Get the address of EbMode in vba6.dll
      If sc_pEbMode = 0 Then                                                            'Found?
        sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                      'VB5 perhaps
      End If
    End If
    
    Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                            'Patch the address of this object instance into the static machine code buffer
    
    sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                             'Get the address of the CallWindowsProc function
    sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                             'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lhWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .sCode = sc_aBuf
    .nAddrSub = StrPtr(.sCode)
    '.nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i)           'Mark memory as executable
    'Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                           'Copy the machine code from the static byte array to the code array in sc_aSubData
    
    .hwnd = lhWnd                                                                       'Store the hWnd
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    
    Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                                     'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                                        'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                                        'Patch the relative address of the CallWindowProc api function
  End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lhWnd As Long)
'Parameters:
  'lhWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lhWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    'Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
'Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'  Dim nEntry As Long
'
'  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
'    nMsgCnt = 0                                                                         'Message count is now zero
'    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
'      nEntry = PATCH_05                                                                 'Patch the before table message count location
'    Else                                                                                'Else after
'      nEntry = PATCH_09                                                                 'Patch the after table message count location
'    End If
'    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
'  Else                                                                                  'Else deleteting a specific message
'    Do While nEntry < nMsgCnt                                                           'For each table entry
'      nEntry = nEntry + 1
'      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
'        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
'        Exit Do                                                                         'Bail
'      End If
'    Loop                                                                                'Next entry
'  End If
'End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lhWnd Then                                                             'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function
