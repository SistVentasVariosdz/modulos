VERSION 5.00
Begin VB.UserControl ucTabMDI 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   HasDC           =   0   'False
   ScaleHeight     =   44
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   362
End
Attribute VB_Name = "ucTabMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------
'Autor:         Leandro Ascierto
'Web:           www.leandroascierto.com.ar
'Date:          09/08/2011
'Test:          Windows XP, Window Seven
'para este proyecto se utilizo parte del codigo del ucTabStrip de Raul338 http://www.leandroascierto.com.ar/foro/index.php?topic=1065.0
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
'  ucTabMDI
'------------------------------------------------------------------------------------------------
'== KERNEL32
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function MulDiv Lib "KERNEL32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
'== Gdi32
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'== User32
Private Declare Function EnableWindow Lib "USER32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetClientRect Lib "USER32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageW Lib "USER32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "USER32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "USER32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenuEx Lib "USER32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hwnd As Long, ByVal lptpm As Any) As Long
Private Declare Function GetSystemMenu Lib "USER32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function DestroyIcon Lib "USER32" (ByVal hIcon As Long) As Long
Private Declare Function RedrawWindow Lib "USER32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function DefWindowProc Lib "USER32" Alias "DefWindowProcW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function LoadIcon Lib "USER32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function DrawFrameControl Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
'== Comctl32
Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hImageList As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_SetImageCount Lib "comctl32.dll" (ByVal himl As Long, ByVal uNewCount As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
'== uxtheme
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As Any) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Private Type NMHDR
    hwndFrom As Long
    idfrom   As Long
    code     As Long
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

'TabStrip
Private Type TCHITTESTINFO
    PT          As POINTAPI
    flags       As Long
End Type

Private Type TCITEM
    mask        As Long
    dwState     As Long
    dwStateMask As Long
    pszText     As Long
    cchTextMax  As Long
    iImage      As Long
    lParam      As Long
End Type

Private Const TCM_FIRST             As Long = &H1300
Private Const TCM_GETIMAGELIST      As Long = (TCM_FIRST + 2)
Private Const TCM_SETIMAGELIST      As Long = (TCM_FIRST + 3)
Private Const TCM_GETITEMCOUNT      As Long = (TCM_FIRST + 4)
Private Const TCM_INSERTITEM        As Long = (TCM_FIRST + 7)
Private Const TCM_DELETEITEM        As Long = (TCM_FIRST + 8)
Private Const TCM_DELETEALLITEMS    As Long = (TCM_FIRST + 9)
Private Const TCM_GETITEMRECT       As Long = (TCM_FIRST + 10)
Private Const TCM_GETCURSEL         As Long = (TCM_FIRST + 11)
Private Const TCM_SETCURSEL         As Long = (TCM_FIRST + 12)
Private Const TCM_HITTEST           As Long = (TCM_FIRST + 13)
Private Const TCM_SETITEMEXTRA      As Long = (TCM_FIRST + 14)
Private Const TCM_ADJUSTRECT        As Long = (TCM_FIRST + 40)
Private Const TCM_SETITEMSIZE       As Long = (TCM_FIRST + 41)
Private Const TCM_REMOVEIMAGE       As Long = (TCM_FIRST + 42)
Private Const TCM_SETPADDING        As Long = (TCM_FIRST + 43)
Private Const TCM_GETROWCOUNT       As Long = (TCM_FIRST + 44)
Private Const TCM_GETTOOLTIPS       As Long = (TCM_FIRST + 45)
Private Const TCM_SETTOOLTIPS       As Long = (TCM_FIRST + 46)
Private Const TCM_GETCURFOCUS       As Long = (TCM_FIRST + 47)
Private Const TCM_SETCURFOCUS       As Long = (TCM_FIRST + 48)
Private Const TCM_SETMINTABWIDTH    As Long = (TCM_FIRST + 49)
Private Const TCM_DESELECTALL       As Long = (TCM_FIRST + 50)
Private Const TCM_HIGHLIGHTITEM     As Long = (TCM_FIRST + 51)
Private Const TCM_SETEXTENDEDSTYLE  As Long = (TCM_FIRST + 52)
Private Const TCM_GETEXTENDEDSTYLE  As Long = (TCM_FIRST + 53)
Private Const TCM_GETITEMW          As Long = (TCM_FIRST + 60)
Private Const TCM_SETITEMW          As Long = (TCM_FIRST + 61)
Private Const TCM_INSERTITEMW       As Long = (TCM_FIRST + 62)
' Styles
Private Const TCS_SINGLELINE        As Long = &H0
Private Const TCS_RIGHTJUSTIFY      As Long = &H0
Private Const TCS_TABS              As Long = &H0
Private Const TCS_SCROLLOPPOSITE    As Long = &H1
Private Const TCS_RIGHT             As Long = &H2
Private Const TCS_BOTTOM            As Long = &H2
Private Const TCS_MULTISELECT       As Long = &H4
Private Const TCS_FLATBUTTONS       As Long = &H8
Private Const TCS_FORCEICONLEFT     As Long = &H10
Private Const TCS_FORCELABELLEFT    As Long = &H20
Private Const TCS_HOTTRACK          As Long = &H40
Private Const TCS_VERTICAL          As Long = &H80
Private Const TCS_BUTTONS           As Long = &H100
Private Const TCS_MULTILINE         As Long = &H200
Private Const TCS_FIXEDWIDTH        As Long = &H400
Private Const TCS_RAGGEDRIGHT       As Long = &H800
Private Const TCS_FOCUSNEVER        As Long = &H8000
Private Const TCS_FOCUSONBUTTONDOWN As Long = &H1000
Private Const TCS_OWNERDRAWFIXED    As Long = &H2000
Private Const TCS_TOOLTIPS          As Long = &H4000
' Ex-Styles
Private Const TCS_EX_FLATSEPARATORS As Long = &H1
Private Const TCS_EX_REGISTERDROP   As Long = &H2
' HitTest
Private Const TCHT_ONITEMICON       As Long = &H2
Private Const TCHT_ONITEMLABEL      As Long = &H4
Private Const TCHT_NOWHERE          As Long = &H1
Private Const TCHT_ONITEM           As Long = (TCHT_ONITEMICON Or TCHT_ONITEMLABEL)
' Item Flags
Private Const TCIF_IMAGE            As Long = &H2
Private Const TCIF_PARAM            As Long = &H8
Private Const TCIF_RTLREADING       As Long = &H4
Private Const TCIF_STATE            As Long = &H10
Private Const TCIF_TEXT             As Long = &H1
' Item States
Private Const TCIS_BUTTONPRESSED    As Long = &H1
Private Const TCIS_HIGHLIGHTED      As Long = &H2
' Notifications
Private Const TCN_FIRST             As Long = -550
Private Const TCN_SELCHANGE         As Long = (TCN_FIRST - 1)
Private Const TCN_SELCHANGING       As Long = (TCN_FIRST - 2)
Private Const TCN_FOCUSCHANGE       As Long = (TCN_FIRST - 4)

Private Const WC_TABCONTROL         As String = "SysTabControl32"

'----------------------
Private Const ILC_MASK          As Long = &H1
Private Const ILC_COLOR32       As Long = &H20
Private Const ILD_TRANSPARENT   As Long = &H1

Private Const GWL_EXSTYLE           As Long = -20
Private Const GWL_STYLE             As Long = (-16)

Private Const WS_CHILD              As Long = &H40000000
Private Const WS_CLIPCHILDREN       As Long = &H2000000
Private Const WS_CLIPSIBLINGS       As Long = &H4000000
Private Const WS_OVERLAPPED         As Long = &H0&
Private Const WS_VISIBLE            As Long = &H10000000
Private Const WS_TABS               As Long = (WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_OVERLAPPED Or WS_VISIBLE Or WS_CHILD)

Private Const WS_EX_MDICHILD        As Long = &H40&
Private Const WS_EX_APPWINDOW       As Long = &H40000

Private Const WM_DESTROY            As Long = &H2
Private Const WM_KILLFOCUS          As Long = &H8
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
Private Const WM_SIZE               As Long = &H5
Private Const WM_SETFONT            As Long = &H30
Private Const WM_NOTIFY             As Long = &H4E
Private Const WM_GETFONT            As Long = &H31
Private Const WM_SETICON            As Long = &H80
Private Const WM_SETTEXT            As Long = &HC
Private Const WM_NCUAHDRAWCAPTION   As Long = &HAE
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_LBUTTONDOWN        As Long = &H201

Private Const NM_FIRST              As Long = 0
Private Const NM_RCLICK             As Long = (NM_FIRST - 5)
Private Const NM_CLICK              As Long = (NM_FIRST - 2)
Private Const NM_LDOWN              As Long = (NM_FIRST - 20)

Private Const DFC_CAPTION           As Long = 1
Private Const DFCS_PUSHED           As Long = &H200
Private Const DFCS_CAPTIONCLOSE     As Long = &H0

Private Const SC_RESTORE            As Long = &HF120&
Private Const SC_MAXIMIZE           As Long = &HF030&
Private Const SC_MINIMIZE           As Long = &HF020&
Private Const SC_NEXTWINDOW         As Long = &HF040&
Private Const SC_CLOSE              As Long = &HF060&

Private Const TPM_RETURNCMD         As Long = &H100&
Private Const TPM_RIGHTALIGN        As Long = &H8&
Private Const TPM_BOTTOMALIGN       As Long = &H20&

Private Const DT_SINGLELINE         As Long = &H20
Private Const DT_VCENTER            As Long = &H4
Private Const DT_WORD_ELLIPSIS      As Long = &H40000

Private Const IDI_APPLICATION       As Long = 32512&
Private Const GA_ROOT               As Long = 2

Private Const LOGPIXELSY            As Long = 90
Private Const FW_NORMAL             As Long = 400
Private Const FW_BOLD               As Long = 700

Public Enum eClsBtnStyle
    MIN_BUTTON = 0
    MDI_BUTTON = 1
End Enum
    
Public Event Resize()

Private WithEvents m_oFont          As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private hMDIClient                  As Long
Private m_bInitialized              As Boolean
Private m_hImageList                As Long
Private m_hFont                     As Long
Private m_IconSize                  As Long
Private m_MaxLen                    As Long
Private m_MinTabsWidth              As Long
Private hTabs                       As Long
Private m_LastTab                   As Long
Private m_CloseButtonVisible        As Boolean
Private m_ShowMenu                  As Boolean
Private m_MultiLine                 As Boolean
Private m_CloseButtonStyle          As eClsBtnStyle
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

    On Error Resume Next
  
    Dim uNMHDR  As NMHDR
    Dim PT As POINTAPI
    Dim Index As Long

    Select Case lng_hWnd

        Case UserControl.hwnd 'UserControl

            Select Case uMsg

                Case WM_NOTIFY
                    
                    Call CopyMemory(uNMHDR, ByVal lParam, Len(uNMHDR))
                    
                    Select Case uNMHDR.code
                     
                        Case TCN_SELCHANGE
                            pvTabSelect pvSelectedItem
                            
                        Case TCN_SELCHANGING
                          
                            If m_CloseButtonVisible Then
                                Index = isCurosrInBtnClose
                                If Index > -1 Then
                                    lReturn = 1
                                    bHandled = True
                                    RedrawWindow hTabs, ByVal 0&, ByVal 0&, &H1
                                End If
                            End If
                        Case NM_CLICK
                            If m_CloseButtonVisible Then
                                Index = isCurosrInBtnClose
                                If Index > -1 Then
                                    SendMessage pvItemParam(Index), WM_SYSCOMMAND, SC_CLOSE, ByVal 0&
                                End If
                                RedrawWindow hTabs, ByVal 0&, ByVal 0&, &H1
                            End If
                            
                        Case NM_RCLICK
                            If m_ShowMenu Then
                                GetCursorPos PT
                                ScreenToClient hTabs, PT
                                Index = pvHitTest(PT.X, PT.Y)
                                If Index > -1 Then
                                    pvToolbarButtonRightClick Index
                                End If
                            End If
                    End Select

                Case WM_SIZE
                     pvAutoSize
                     Call MoveWindow(hTabs, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1)
            
            End Select

        Case hMDIClient 'MDICLIENT
   
            Select Case uMsg
            
                Case WM_KILLFOCUS
                
                    If GetProp(wParam, "SubClass") = 0 Then
                        If pvIsShowInTaskBar(wParam) Then
                            SetProp wParam, "SubClass", 1
                            Call Subclass_Start(wParam)
                            Call Subclass_AddMsg(wParam, WM_SHOWWINDOW, MSG_AFTER)
                            Call Subclass_AddMsg(wParam, WM_DESTROY, MSG_BEFORE)
                            Call Subclass_AddMsg(wParam, WM_SYSCOMMAND, MSG_BEFORE_AND_AFTER)
                            Call Subclass_AddMsg(wParam, WM_MDIACTIVATE, MSG_AFTER)
                            Call Subclass_AddMsg(wParam, WM_SETICON, MSG_AFTER)
                            Call Subclass_AddMsg(wParam, WM_STYLECHANGED, MSG_AFTER)
                            Call Subclass_AddMsg(wParam, WM_SETTEXT, MSG_AFTER)
                            Call Subclass_AddMsg(wParam, WM_NCUAHDRAWCAPTION, MSG_AFTER)
    
                            Call pvAddWindow(wParam)
                        End If
                    End If

                    pvFindActive wParam

            End Select
            
        Case hTabs 'TabControl
             If Not m_CloseButtonVisible Then Exit Sub
             
            Select Case uMsg
                
                Case WM_PAINT
                    DrawCloseButtons
                
                Case WM_MOUSEMOVE
                    Index = isCurosrInBtnClose
                    If Index <> -1 Then
                        If m_LastTab <> Index Then
                            RedrawWindow hTabs, ByVal 0&, ByVal 0&, &H1
                            m_LastTab = Index
                        End If
                    Else
                        If m_LastTab <> -1 Then
                            RedrawWindow hTabs, ByVal 0&, ByVal 0&, &H1
                            m_LastTab = -1
                        End If
                    End If

                Case WM_LBUTTONDOWN
                    RedrawWindow hTabs, ByVal 0&, ByVal 0&, &H1

            End Select
            
        Case Else 'Child Windows

            Select Case uMsg
            
                Case WM_DESTROY
                    Subclass_Stop lng_hWnd
                    
                Case WM_SHOWWINDOW
                    If wParam = 1 Then
                        If Not pvExistWindowTab(lng_hWnd) Then
                            Call pvAddWindow(lng_hWnd)
                        End If
                    Else
                        Call pvRemoveWindow(lng_hWnd)
                    End If
                    
                Case WM_MDIACTIVATE
                    Call pvFindActive(lParam)
                    
                Case WM_SETICON
                    Call pvChangeIcon(lng_hWnd)
                
                Case WM_STYLECHANGED, WM_NCUAHDRAWCAPTION, WM_SETTEXT
                    Call pvChangeCaption(lng_hWnd)

            End Select
        
    End Select
End Sub

'========================================================================================
'Public Function
'========================================================================================
Public Function Initialize(ByVal hwndMDI As Long) As Boolean

        If (m_bInitialized = False) Then

            hMDIClient = FindWindowEx(hwndMDI, ByVal 0&, "MDIClient", vbNullString)
  
            If hMDIClient Then
                If m_IconSize = 0 Then m_IconSize = 16

                If pvCreateTabControl(m_IconSize) Then

                    m_bInitialized = True

                    '-- Subclass
                    Call Subclass_Start(UserControl.hwnd)
                    Call Subclass_AddMsg(UserControl.hwnd, WM_SIZE, MSG_BEFORE)
                    Call Subclass_AddMsg(UserControl.hwnd, WM_NOTIFY, MSG_BEFORE)
                    
                    Call Subclass_Start(hMDIClient)
                    Call Subclass_AddMsg(hMDIClient, WM_KILLFOCUS, MSG_AFTER)

                    Call Subclass_Start(hTabs)
                    Call Subclass_AddMsg(hTabs, WM_MOUSEMOVE, MSG_BEFORE)
                    Call Subclass_AddMsg(hTabs, WM_LBUTTONDOWN, MSG_BEFORE)
                    
                    Call Subclass_AddMsg(hTabs, WM_PAINT, MSG_AFTER)
    
                    Initialize = True

                End If
            End If
        End If
End Function

'========================================================================================
' Public Property
'========================================================================================
Public Property Get hwndTabControl() As Long
    hwndTabControl = hTabs
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
    If hTabs Then Call EnableWindow(hTabs, New_Enabled)
End Property

Public Property Get IconSize() As Long
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_Size As Long)
    Dim i As Long
    Dim hIcon As Long
    Dim lCount As Long
    
    m_IconSize = New_Size
    
    If hTabs Then
        If m_hImageList Then ImageList_Destroy m_hImageList
    
        m_hImageList = ImageList_Create(m_IconSize, m_IconSize, ILC_MASK Or ILC_COLOR32, 0, 0)
        
        If m_hImageList Then
            Call SendMessage(hTabs, TCM_SETIMAGELIST, 0&, ByVal m_hImageList)

            For i = 0 To pvTabsCount - 1
        
                hIcon = pvGetWindowIcon(pvItemParam(i))
                
                If hIcon = 0 Then
                    hIcon = LoadIcon(0&, IDI_APPLICATION)
                    lCount = ImageList_AddIcon(m_hImageList, hIcon)
                    DestroyIcon hIcon
                Else
                    lCount = ImageList_AddIcon(m_hImageList, hIcon)
                End If
    
                pvItemImage(i) = lCount
            
            Next
        End If
    End If
    pvAutoSize
End Property

Public Property Get Font() As StdFont
    Set Font = m_oFont
End Property

Public Property Set Font(ByVal New_Font As StdFont)

    Dim uLF   As LOGFONT
    Dim lChar As Long
    
    Set m_oFont = New_Font
    Set UserControl.Font = New_Font
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
        
    If hTabs Then Call SendMessage(hTabs, WM_SETFONT, m_hFont, ByVal 1&)
    pvAutoSize
End Property

Public Property Get Multiline() As Boolean
    Multiline = m_MultiLine
End Property

Public Property Let Multiline(ByVal NewValue As Boolean)
    m_MultiLine = NewValue
    If hTabs Then
        Call SetWindowLongA(hTabs, GWL_STYLE, (GetWindowLong(hTabs, GWL_STYLE) And Not TCS_MULTILINE) Or (NewValue And TCS_MULTILINE))
        pvAutoSize
    End If
    PropertyChanged "MultiLine"
End Property

Public Property Let MinTabWidth(ByVal newMinWidth As Long)
    m_MinTabsWidth = newMinWidth
    Call SendMessageW(hTabs, TCM_SETMINTABWIDTH, 0, ByVal newMinWidth)
    pvAutoSize
    PropertyChanged "MinTabsWidth"
End Property

Public Property Get MinTabWidth() As Long
    MinTabWidth = m_MinTabsWidth
End Property

Public Property Let MaxLen(ByVal NewMaxLen As Long)
    m_MaxLen = NewMaxLen
    RefreshCaptions
    PropertyChanged "MaxLen"
End Property

Public Property Get MaxLen() As Long
    MaxLen = m_MaxLen
End Property

Public Property Let CloseButtonVisible(ByVal NewValue As Boolean)
    m_CloseButtonVisible = NewValue
    RefreshCaptions
    PropertyChanged "CloseButtonVisible"
End Property

Public Property Get CloseButtonVisible() As Boolean
    CloseButtonVisible = m_CloseButtonVisible
End Property

Public Property Let CloseButtonStyle(ByVal NewStyle As eClsBtnStyle)
    m_CloseButtonStyle = NewStyle
    If hTabs Then RedrawWindow hTabs, ByVal 0&, ByVal 0&, &H1
    PropertyChanged "CloseButtonStyle"
End Property

Public Property Get CloseButtonStyle() As eClsBtnStyle
    CloseButtonStyle = m_CloseButtonStyle
End Property

Public Property Let ShowMenu(ByVal NewValue As Boolean)
    m_ShowMenu = NewValue
    PropertyChanged "ShowMenu"
End Property

Public Property Get ShowMenu() As Boolean
    ShowMenu = m_ShowMenu
End Property

'========================================================================================
' Usercontrol
'========================================================================================
Private Sub UserControl_Initialize()
    Call InitCommonControls
    Set m_oFont = New StdFont
    m_IconSize = 16
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    m_IconSize = 16
    m_CloseButtonVisible = True
    m_ShowMenu = True
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
        Call pvDestroyTabControl
    End If
errH:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_IconSize = .ReadProperty("IconSize", 32)
        m_MinTabsWidth = .ReadProperty("MinTabsWidth", 0)
        m_MaxLen = .ReadProperty("MaxLen", 0)
        m_CloseButtonVisible = .ReadProperty("CloseButtonVisible", True)
        m_ShowMenu = .ReadProperty("ShowMenu", True)
        m_MultiLine = .ReadProperty("MultiLine", False)
        m_CloseButtonStyle = .ReadProperty("CloseButtonStyle", MIN_BUTTON)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set Me.Font = .ReadProperty("Font", Ambient.Font)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("IconSize", m_IconSize, 32)
        Call .WriteProperty("MinTabsWidth", m_MinTabsWidth, 0)
        Call .WriteProperty("MaxLen", m_MaxLen, 0)
        Call .WriteProperty("CloseButtonVisible", m_CloseButtonVisible, True)
        Call .WriteProperty("ShowMenu", m_ShowMenu, True)
        Call .WriteProperty("MultiLine", m_MultiLine, False)
        Call .WriteProperty("CloseButtonStyle", m_CloseButtonStyle, MIN_BUTTON)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Font", m_oFont, Ambient.Font)
    End With
End Sub

'========================================================================================
' Private Sub Funtion and Property
'========================================================================================
Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Set Font = m_oFont
End Sub

Private Sub pvAutoSize()
    Dim Rec As RECT

    If pvTabsCount Then
        SendMessage hTabs, TCM_GETITEMRECT, pvSelectedItem, Rec
        UserControl.Height = (Rec.Bottom + 4) * Screen.TwipsPerPixelY
    Else
        UserControl.Height = 1
    End If
End Sub

Private Sub DrawCloseButtons()
    Dim i As Long
    Dim Rec As RECT
    Dim hdc As Long
    Dim lTabW As Long
    Dim lTabH As Long
    Dim iSeleted As Long
    Dim PT As POINTAPI
    Dim iState As Long
    Dim hTheme  As Long
    
    iSeleted = pvSelectedItem
    hdc = GetDC(hTabs)
    
    hTheme = OpenThemeData(0&, StrPtr("Window"))
    
    For i = 0 To pvTabsCount - 1
        SendMessage hTabs, TCM_GETITEMRECT, i, Rec
        lTabW = Rec.Right - Rec.Left
        lTabH = Rec.Bottom - Rec.Top
        If isCurosrInBtnClose = i Then
            If GetKeyState(1) < 0 Then
                iState = 3
            Else
                iState = 2
            End If
        Else
            iState = 1
        End If
        If i = iSeleted Then
            DrawTheme hdc, hTheme, 19& + m_CloseButtonStyle, iState, Rec.Right - 16, Rec.Top + (lTabH / 2) - 8, 13, 13
        Else
            DrawTheme hdc, hTheme, 19& + m_CloseButtonStyle, iState, Rec.Right - 18, Rec.Top + (lTabH / 2) - 6, 13, 13
        End If
    Next
    
    If hTheme Then CloseThemeData hTheme
    ReleaseDC 0&, hdc
End Sub

Private Function DrawTheme(hdc As Long, hTheme As Long, PartId As Long, StateId As Long, Left As Long, Top As Long, Width As Long, Height As Long) As Boolean
    Dim Rec As RECT
    
    With Rec
       .Left = Left
       .Top = Top
       .Right = Left + Width
       .Bottom = Top + Height
    End With

    If (hTheme) Then
        DrawTheme = DrawThemeBackground(hTheme, hdc, PartId, StateId, Rec, ByVal 0&) = 0
    Else
        DrawFrameControl hdc, Rec, DFC_CAPTION, DFCS_CAPTIONCLOSE Or IIf(StateId = 3, DFCS_PUSHED, 0)
    End If
    
End Function

Private Function isCurosrInBtnClose() As Long
    Dim Index As Long
    Dim PT As POINTAPI
    GetCursorPos PT
    ScreenToClient hTabs, PT
    Index = pvHitTest(PT.X, PT.Y)
    If Index > -1 Then
        If isHiTestBtnClose(Index, PT.X, PT.Y) Then
            isCurosrInBtnClose = Index
        Else
            isCurosrInBtnClose = -1
        End If
    Else
        isCurosrInBtnClose = -1
    End If
End Function

Private Function isHiTestBtnClose(ByVal Index As Long, X As Long, Y As Long) As Boolean
    Dim Rec As RECT
    Dim RectBtn As RECT
    SendMessage hTabs, TCM_GETITEMRECT, Index, Rec
    If Index = pvSelectedItem Then
        With RectBtn
            .Left = Rec.Right - 16
            .Top = Rec.Top + ((Rec.Bottom - Rec.Top) / 2) - 8
            .Right = .Left + 13
            .Bottom = .Top + 13
        End With
    Else
        With RectBtn
            .Left = Rec.Right - 18
            .Top = Rec.Top + ((Rec.Bottom - Rec.Top) / 2) - 6
            .Right = .Left + 13
            .Bottom = .Top + 13
        End With
    End If
    isHiTestBtnClose = PtInRect(RectBtn, X, Y)
End Function

Private Sub pvTabSelect(ByVal Index As Long)

    Dim handle As Long
    Dim hActive As Long

    hActive = SendMessage(hMDIClient, WM_MDIGETACTIVE, 0&, ByVal 0&)

    handle = pvItemParam(Index)
    
    If IsIconic(handle) Then
        If IsZoomed(hActive) Then
            SendMessage handle, WM_SYSCOMMAND, SC_MAXIMIZE, ByVal 0&
        Else
            SendMessage handle, WM_SYSCOMMAND, SC_RESTORE, ByVal 0&
        End If
    Else
        BringWindowToTop handle
    End If

End Sub

Private Sub pvToolbarButtonRightClick(ByVal Index As Long)
    Dim handle As Long
    Dim lRet As Long
    Dim PT As POINTAPI
    Dim i As Long
   
    handle = pvItemParam(Index)
    'BringWindowToTop handle
    GetCursorPos PT
    lRet = TrackPopupMenuEx(GetSystemMenu(handle, False), TPM_RETURNCMD, PT.X, PT.Y, UserControl.hwnd, ByVal 0&)
    If lRet Then SendMessage handle, WM_SYSCOMMAND, lRet, ByVal 0&
End Sub

Private Sub RefreshCaptions()
    Dim i As Long
    For i = 0 To pvTabsCount - 1
        pvChangeCaption pvItemParam(i)
    Next
End Sub

Private Sub pvChangeCaption(hwnd As Long)
    Dim sCaption As String
    Dim i As Long
    
    sCaption = pvGetWindowTextW(hwnd)
    
    If m_MaxLen > 0 Then
        sCaption = GetEndEllipsisText(sCaption)
    End If
    
    If m_CloseButtonVisible Then
        sCaption = sCaption & Space(48 / m_oFont.SIZE)
    End If

    For i = 0 To pvTabsCount - 1
        If pvItemParam(i) = hwnd Then
            pvItemText(i) = sCaption
            Exit For
        End If
    Next
End Sub

Private Function GetEndEllipsisText(ByVal sText As String) As String
        If Len(sText) > m_MaxLen Then
            GetEndEllipsisText = Left$(sText, m_MaxLen) & "..."
        Else
            GetEndEllipsisText = sText
        End If
End Function

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
    
    For i = 0 To pvTabsCount - 1
        If pvItemParam(i) = hwnd Then
            pvItemImage(i) = lCount
            Exit For
        End If
    Next
End Sub

Private Function pvExistWindowTab(hwnd As Long) As Boolean
    Dim i As Long

    For i = 0 To pvTabsCount - 1
        If pvItemParam(i) = hwnd Then
            pvExistWindowTab = True
            Exit Function
        End If
    Next
End Function

Private Sub pvAddWindow(hwnd As Long)
    Dim sCaption As String
    Dim hIcon As Long
    Dim lCount As Long
    
    sCaption = pvGetWindowTextW(hwnd)
    
    If m_MaxLen > 0 Then
        sCaption = GetEndEllipsisText(sCaption)
    End If
    
    If m_CloseButtonVisible Then
        sCaption = sCaption & Space(48 / m_oFont.SIZE)
    End If
    hIcon = pvGetWindowIcon(hwnd)
    
    If hIcon = 0 Then
        hIcon = LoadIcon(0&, IDI_APPLICATION)
        lCount = ImageList_AddIcon(m_hImageList, hIcon)
        DestroyIcon hIcon
    Else
        lCount = ImageList_AddIcon(m_hImageList, hIcon)
    End If

    pvAddTab sCaption, lCount, hwnd
    pvAutoSize
End Sub

Private Sub pvRemoveWindow(hwnd As Long)
    Dim i As Long
    Dim lCount As Long
    m_LastTab = -1
    lCount = pvTabsCount
    
    For i = 0 To lCount - 1
        If pvItemParam(i) = hwnd Then
            pvRemoveTab i
            Exit For
        End If
    Next

    If lCount = 1 Then Call ImageList_SetImageCount(m_hImageList, 0&)
    pvAutoSize
End Sub

Private Function pvFindActive(hActive As Long)
    Dim i As Long

    For i = 0 To pvTabsCount - 1
        If pvItemParam(i) = hActive Then
            Call SendMessageW(hTabs, TCM_SETCURSEL, i, ByVal 0&)
            Exit For
        End If
    Next
End Function

Private Function pvCreateTabControl(ByVal ImageSize As Long) As Boolean
    Dim lStyle As Long
    
    If m_MultiLine Then lStyle = TCS_MULTILINE

    hTabs = CreateWindowEx(0&, WC_TABCONTROL, vbNullString, WS_TABS Or lStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hwnd, 0, App.hInstance, ByVal 0&)

    If hTabs Then
        
        m_hImageList = ImageList_Create(ImageSize, ImageSize, ILC_MASK Or ILC_COLOR32, 0&, 0&)
        
        If m_hImageList Then Call SendMessage(hTabs, TCM_SETIMAGELIST, 0&, ByVal m_hImageList)

        SendMessage hTabs, WM_SETFONT, m_hFont, ByVal 0&
        SendMessage hTabs, TCM_SETMINTABWIDTH, 0, ByVal m_MinTabsWidth

        pvCreateTabControl = True
    End If
End Function

Private Sub pvDestroyTabControl()
    If (hTabs) Then
        If (DestroyWindow(hTabs)) Then
            hTabs = 0
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

Private Property Get pvTabsCount() As Long
    pvTabsCount = SendMessageW(hTabs, TCM_GETITEMCOUNT, 0, ByVal 0)
End Property

Private Property Get pvItemImage(ByVal Index As Long) As Long
    Dim sTabStrip As TCITEM
    sTabStrip.mask = TCIF_IMAGE
    Call SendMessageW(hTabs, TCM_GETITEMW, Index, sTabStrip)
    pvItemImage = sTabStrip.iImage
End Property

Private Property Let pvItemImage(ByVal Index As Long, ByVal iImage As Long)
    Dim sTabSrip As TCITEM
    sTabSrip.mask = TCIF_IMAGE
    sTabSrip.iImage = iImage
    Call SendMessageW(hTabs, TCM_SETITEMW, Index, sTabSrip)
End Property

Private Property Get pvItemParam(ByVal Index As Long) As Long
    Dim sTabStrip As TCITEM
    sTabStrip.mask = TCIF_PARAM
    Call SendMessageW(hTabs, TCM_GETITEMW, Index, sTabStrip)
    pvItemParam = sTabStrip.lParam
End Property

Private Property Let pvItemParam(ByVal Index As Long, ByVal lParam As Long)
    Dim sTabSrip As TCITEM
    sTabSrip.mask = TCIF_PARAM
    sTabSrip.lParam = lParam
    Call SendMessageW(hTabs, TCM_SETITEMW, Index, sTabSrip)
End Property

Private Property Get pvItemText(ByVal Index As Long) As String
    If hTabs Then
        Dim sTabStrip As TCITEM
        Dim sText As String
        sText = String(255, 0)
        sTabStrip.mask = TCIF_TEXT
        sTabStrip.cchTextMax = 255
        sTabStrip.pszText = StrPtr(sText)
        If SendMessageW(hTabs, TCM_GETITEMW, Index, sTabStrip) Then pvItemText = Left$(sText, InStr(sText, vbNullChar) - 1)
    End If
End Property

Private Property Let pvItemText(ByVal Index As Long, ByVal text As String)
    If hTabs Then
        Dim sTabSrip As TCITEM
        sTabSrip.mask = TCIF_TEXT
        sTabSrip.pszText = StrPtr(text)
        Call SendMessageW(hTabs, TCM_SETITEMW, Index, sTabSrip)
    End If
End Property

Private Property Get pvSelectedItem() As Long
    If hTabs Then pvSelectedItem = SendMessageW(hTabs, TCM_GETCURSEL, 0, 0)
End Property

Private Property Let pvSelectedItem(ByVal Index As Long)
    If hTabs <> 0 Then Call SendMessageW(hTabs, TCM_SETCURSEL, Index, ByVal 0)
End Property

Private Sub pvAddTab(sCaption As String, ImageIndex As Long, ItemData As Long)

    Dim lCount As Long
    Dim sTabSrip As TCITEM
    
    lCount = SendMessageW(hTabs, TCM_GETITEMCOUNT, 0&, ByVal 0&)

    With sTabSrip
        .mask = TCIF_TEXT Or TCIF_IMAGE Or TCIF_PARAM
        .iImage = ImageIndex
        .lParam = ItemData
        .pszText = StrPtr(sCaption)
    End With
    
    Call SendMessageW(hTabs, TCM_INSERTITEMW, lCount, sTabSrip)

End Sub

Private Function pvHitTest(ByVal X As Single, ByVal Y As Single) As Long
    If hTabs Then
        Dim HT As TCHITTESTINFO
        HT.PT.X = X: HT.PT.Y = Y
        pvHitTest = SendMessageW(hTabs, TCM_HITTEST, 0, HT)
    End If
End Function

Public Function pvRemoveTab(ByVal Index As Long) As Boolean
    pvRemoveTab = SendMessageW(hTabs, TCM_DELETEITEM, Index, ByVal 0&)
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
