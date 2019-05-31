VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB03_WINEVE_TOOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' The NMHDR structure contains information about
' a notification message. The pointer  to this
' structure is specified as the lParam member of
' the WM_NOTIFY message.
'
Public Type NMHDR
    hwndFrom As Long
    idFrom   As Long
    code     As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const WM_USER = &H400
Private Const TOOLTIPS_CLASS = "tooltips_class32"
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_NOPREFIX = &H2

#Const WIN32_IE = &H400

Public Type TOOLINFO
    cbSize   As Long
    uFlags   As TT_Flags
    hwnd     As Long
    uId      As Long
    RECT     As RECT
    hinst    As Long
    lpszText As String
    #If (WIN32_IE >= &H300) Then
        lParam   As Long
    #End If
End Type

Public Enum TT_Flags
    TTF_IDISHWND = &H1
    TTF_CENTERTIP = &H2
    TTF_RTLREADING = &H4
    TTF_SUBCLASS = &H10
    #If (WIN32_IE >= &H300) Then
        TTF_TRACK = &H20
        TTF_ABSOLUTE = &H80
        TTF_TRANSPARENT = &H100
        TTF_DI_SETITEM = &H8000&
    #End If
End Enum

Public Enum TT_DelayTime
    TTDT_AUTOMATIC = 0
    TTDT_RESHOW = 1
    TTDT_AUTOPOP = 2
    TTDT_INITIAL = 3
End Enum

Public Enum ttDelayTimeConstants
    ttDelayDefault = TTDT_AUTOMATIC '= 0
    ttDelayInitial = TTDT_INITIAL '= 3
    ttDelayShow = TTDT_AUTOPOP '= 2
    ttDelayReshow = TTDT_RESHOW '= 1
    ttDelayMask = 3
End Enum

Public Enum ttMarginConstants
    ttMarginLeft = 0
    ttMarginTop = 1
    ttMarginRight = 2
    ttMarginBottom = 3
End Enum

Public Type TTHITTESTINFO
    hwnd As Long
    pt   As POINTAPI
    ti   As TOOLINFO
End Type

Public Enum TT_Msgs
    TTM_ACTIVATE = (WM_USER + 1)
    TTM_SETDELAYTIME = (WM_USER + 3)
    TTM_RELAYEVENT = (WM_USER + 7)
    TTM_GETTOOLCOUNT = (WM_USER + 13)
    TTM_WINDOWFROMPOINT = (WM_USER + 16)
    #If UNICODE Then
        TTM_ADDTOOL = (WM_USER + 50)
        TTM_DELTOOL = (WM_USER + 51)
        TTM_NEWTOOLRECT = (WM_USER + 52)
        TTM_GETTOOLINFO = (WM_USER + 53)
        TTM_SETTOOLINFO = (WM_USER + 54)
        TTM_HITTEST = (WM_USER + 55)
        TTM_GETTEXT = (WM_USER + 56)
        TTM_UPDATETIPTEXT = (WM_USER + 57)
        TTM_ENUMTOOLS = (WM_USER + 58)
        TTM_GETCURRENTTOOL = (WM_USER + 59)
    #Else
        TTM_ADDTOOL = (WM_USER + 4)
        TTM_DELTOOL = (WM_USER + 5)
        TTM_NEWTOOLRECT = (WM_USER + 6)
        TTM_GETTOOLINFO = (WM_USER + 8)
        TTM_SETTOOLINFO = (WM_USER + 9)
        TTM_HITTEST = (WM_USER + 10)
        TTM_GETTEXT = (WM_USER + 11)
        TTM_UPDATETIPTEXT = (WM_USER + 12)
        TTM_ENUMTOOLS = (WM_USER + 14)
        TTM_GETCURRENTTOOL = (WM_USER + 15)
    #End If

    #If (WIN32_IE >= &H300) Then
        TTM_TRACKACTIVATE = (WM_USER + 17)
        TTM_TRACKPOSITION = (WM_USER + 18)
        TTM_SETTIPBKCOLOR = (WM_USER + 19)
        TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
        TTM_GETDELAYTIME = (WM_USER + 21)
        TTM_GETTIPBKCOLOR = (WM_USER + 22)
        TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
        TTM_SETMAXTIPWIDTH = (WM_USER + 24)
        TTM_GETMAXTIPWIDTH = (WM_USER + 25)
        TTM_SETMARGIN = (WM_USER + 26)
        TTM_GETMARGIN = (WM_USER + 27)
        TTM_POP = (WM_USER + 28)
    #End If
    
    #If (WIN32_IE >= &H400) Then
        TTM_UPDATE = (WM_USER + 29)
    #End If
End Enum

Public Enum TT_Notifications
    TTN_FIRST = -520&
    TTN_LAST = -549&
    #If UNICODE Then
        TTN_NEEDTEXT = (TTN_FIRST - 10)
    #Else
        TTN_NEEDTEXT = (TTN_FIRST - 0)
    #End If
    TTN_SHOW = (TTN_FIRST - 1)
    TTN_POP = (TTN_FIRST - 2)
End Enum

Private Type NMTTDISPINFO
    hdr      As NMHDR
    lpszText As Long
    #If UNICODE Then
        szText As String * 160
    #Else
        szText As String * 80
    #End If
    hinst  As Long
    uFlags As Long
    #If (WIN32_IE >= &H300) Then
        lParam As Long
    #End If
End Type

'
' Exported by Comctl32.dll >= v4.00.950
' Ensures that the common control dynamic
' link library (DLL) is loaded.
'
' NOTE: API replaced by InitCommonControlsEx
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Declare Function SendMessageT _
        Lib "user32" _
        Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function CreateWindowEx _
        Lib "user32" _
        Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, _
                                 ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, _
                                 ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Private Declare Function DestroyWindow _
        Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub MoveMemory _
        Lib "kernel32" _
        Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)



'
' Defaults:
' DelayInitial   = 500  (1/2 sec)
' DelayAutoPopup = 5000  (5 secs)
' DelayReshow    = 100 (1/10 sec)
' MaxTipWidth    = 0
' all Margins    = 0

Private mnlgHwndTT As Long
Private mnlgMaxTip As Long


Public Function Create(ByRef xForm As Object) As Boolean
    If (mnlgHwndTT = 0) Then
        Call InitCommonControls
        '
        ' The hwndParent param lets the tooltip window
        ' be owned by the specified form and be destroyed
        ' along with it. We'll cleanup in Class_Terminate anyway.
        ' No WS_EX_TOPMOST or TTS_ALWAYSTIP per Win95 UI rules.
        '
        mnlgHwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP, 0, 0, 0, 0, xForm.hwnd, 0, App.hInstance, ByVal 0)
    End If
    Create = CBool(mnlgHwndTT)
End Function

Private Sub Class_Terminate()
    If mnlgHwndTT > 0 Then Call DestroyWindow(mnlgHwndTT)
End Sub

Public Function AddTool(ByRef xControl As Object, Optional ByVal strText As String) As Boolean
    Dim ti As TOOLINFO
  
    If (mnlgHwndTT = 0) Then Exit Function
    
    If (fGetToolInfo(xControl.hwnd, ti) = False) Then
        With ti
            .cbSize = Len(ti)
            '
            ' TTF_IDISHWND must be set to tell the tooltip
            ' control to retrieve the control's rect from
            ' it's hWnd specified in uId.
            '
            .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
            .hwnd = xControl.Container.hwnd
            .uId = xControl.hwnd
            
            If Len(strText) > 0 Then
                .lpszText = strText
'            Else
'                .lpszText = "Tool" & ToolCount + 1
            End If
            '
            ' Maintain the maximun tip text
            ' length for fGetToolInfo.
            '
            mnlgMaxTip = fMax(mnlgMaxTip, Len(.lpszText) + 1)
        End With
        '
        ' Returns 1 on success, 0 on failure
        '
        AddTool = SendMessageT(mnlgHwndTT, TTM_ADDTOOL, 0, ti)
    End If
End Function

Private Function fMax(ByVal lngParm1 As Long, ByVal lngParm2 As Long) As Long
    '
    ' Returns the larger of the two values.
    '
    If lngParm1 > lngParm2 Then
        fMax = lngParm1
    Else
        fMax = lngParm2
    End If
    
End Function

Public Function RemoveTool(ByRef xControl As Object) As Boolean
    Dim ti As TOOLINFO

    If (mnlgHwndTT = 0) Then Exit Function
    
    If fGetToolInfo(xControl.hwnd, ti) Then
        Call SendMessageT(mnlgHwndTT, TTM_DELTOOL, 0, ti)
        RemoveTool = True
    End If
End Function

Public Property Get BackColor() As OLE_COLOR
    If (mnlgHwndTT = 0) Then Exit Property
    
    '
    ' OLE_COLOR is defined in stdole2.tlb
    '
    BackColor = SendMessageT(mnlgHwndTT, TTM_GETTIPBKCOLOR, 0, 0)
End Property

Public Property Let BackColor(clr As OLE_COLOR)
    If (mnlgHwndTT = 0) Then Exit Property
    Call SendMessageT(mnlgHwndTT, TTM_SETTIPBKCOLOR, clr, 0)
End Property

Public Property Get DelayTime(dwType As ttDelayTimeConstants) As Long
    If (mnlgHwndTT = 0) Then Exit Property
    DelayTime = SendMessageT(mnlgHwndTT, TTM_GETDELAYTIME, (dwType And ttDelayMask), 0&)
End Property

Public Property Let DelayTime(dwType As ttDelayTimeConstants, dwMilliSecs As Long)
    If (mnlgHwndTT = 0) Then Exit Property
    Call SendMessageT(mnlgHwndTT, TTM_SETDELAYTIME, (dwType And ttDelayMask), ByVal dwMilliSecs)  ' no rtn val
End Property

Public Property Get ForeColor() As OLE_COLOR
    If (mnlgHwndTT = 0) Then Exit Property
    ForeColor = SendMessageT(mnlgHwndTT, TTM_SETTIPTEXTCOLOR, 0, 0)
End Property

Public Property Let ForeColor(clr As OLE_COLOR)
    If (mnlgHwndTT = 0) Then Exit Property
    Call SendMessageT(mnlgHwndTT, TTM_SETTIPTEXTCOLOR, clr, 0)   ' no rtn val
End Property

Public Property Get hwnd() As Long
    hwnd = mnlgHwndTT
End Property

Public Property Get Margin(dwType As ttMarginConstants) As Long
    Dim rc As RECT

    If (mnlgHwndTT = 0) Then Exit Property
    
    Call SendMessageT(mnlgHwndTT, TTM_GETMARGIN, 0, rc)
    
    Select Case dwType
        Case ttMarginLeft
            Margin = rc.Left
        Case ttMarginTop
            Margin = rc.Top
        Case ttMarginRight
            Margin = rc.Right
        Case ttMarginBottom
            Margin = rc.Bottom
    End Select
End Property

Public Property Let Margin(dwType As ttMarginConstants, cPixels As Long)
    Dim rc As RECT
    
    If (mnlgHwndTT = 0) Then Exit Property
    
    Call SendMessageT(mnlgHwndTT, TTM_GETMARGIN, 0, rc)
    
    Select Case dwType
        Case ttMarginLeft
            rc.Left = cPixels
        Case ttMarginTop
            rc.Top = cPixels
        Case ttMarginRight
            rc.Right = cPixels
        Case ttMarginBottom
            rc.Bottom = cPixels
    End Select
    
    Call SendMessageT(mnlgHwndTT, TTM_SETMARGIN, 0, rc)
End Property

Public Property Get MaxTipWidth() As Long
    '
    ' If MaxTipWidth is -1, there is no word wrapping and
    ' text control characters are printed and not
    ' evaluated (i.e. a vbCrLf shows up as "||")
    '
    If (mnlgHwndTT = 0) Then Exit Property
    MaxTipWidth = fLowWord(SendMessageT(mnlgHwndTT, TTM_GETMAXTIPWIDTH, 0, 0))
End Property

Private Function fLowWord(ByVal lngValue As Long) As Integer
    '
    ' Returns the low-order word from a 32-bit value.
    '
    Call MoveMemory(fLowWord, lngValue, 2)
End Function

Public Property Let MaxTipWidth(ByVal lngWidth As Long)
    '
    ' If MaxTipWidth is -1, there is no word wrapping and
    ' text control characters are printed and not
    ' evaluated (i.e. a vbCrLf shows up as "||")
    '
    If mnlgHwndTT = 0 Then Exit Property
    If lngWidth < 1 Then lngWidth = -1
    Call SendMessageT(mnlgHwndTT, TTM_SETMAXTIPWIDTH, 0, lngWidth)
End Property

Public Property Get ToolCount() As Long
    If (mnlgHwndTT = 0) Then Exit Property
    ToolCount = SendMessageT(mnlgHwndTT, TTM_GETTOOLCOUNT, 0, 0)
End Property

Public Property Get ToolTipHandle() As Long
    ToolTipHandle = mnlgHwndTT
End Property

Public Property Get ToolText(ByRef xControl As Object) As String
    Dim ti As TOOLINFO
  
    If (mnlgHwndTT = 0) Then Exit Property
    If fGetToolInfo(xControl.hwnd, ti, True) Then
        ToolText = fGetStrFromBuffer(ti.lpszText)
    End If
End Property

Private Function fGetStrFromBuffer(ByVal strValue As String) As String
    If InStr(strValue, vbNullChar) Then
        fGetStrFromBuffer = Left$(strValue, InStr(strValue, vbNullChar) - 1)
    Else
        '
        ' If strValue had no null char, the Left$ function
        ' above would rtn a zero length string ("").
        '
        fGetStrFromBuffer = strValue
    End If
End Function

Public Property Let ToolText(ByRef xControl As Object, ByVal strText As String)
    Dim ti As TOOLINFO
  
    If (mnlgHwndTT = 0) Then Exit Property
    
    If fGetToolInfo(xControl.hwnd, ti) Then
        ti.lpszText = strText
        mnlgMaxTip = fMax(mnlgMaxTip, Len(strText) + 1)
        '
        ' The tooltip won't appear for the control
        ' if lpszText is an empty string
        '
        Call SendMessageT(mnlgHwndTT, TTM_UPDATETIPTEXT, 0, ti)
    End If
End Property

Private Function fIsWindow(ByRef xControl As Object) As Boolean
    On Error GoTo ErrorHandler
    fIsWindow = CBool(xControl.hwnd)
ErrorHandler:
End Function

Private Function fGetToolInfo(ByVal lnghwndTool As Long, ti As TOOLINFO, _
                              Optional fGetText As Boolean = False) As Boolean
  
    Dim nItems As Long
    Dim i      As Integer

    ti.cbSize = Len(ti)
    If fGetText Then ti.lpszText = String$(mnlgMaxTip, 0)
    
    nItems = ToolCount
    
    For i = 0 To nItems - 1
        '
        ' Returns 1 on success, 0 on failure.
        '
        If SendMessageT(mnlgHwndTT, TTM_ENUMTOOLS, (i), ti) Then
            If (lnghwndTool = ti.uId) Then
                fGetToolInfo = True
                Exit Function
            End If
        End If
    Next
End Function




