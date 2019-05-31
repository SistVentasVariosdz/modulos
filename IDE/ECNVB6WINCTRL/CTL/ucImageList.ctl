VERSION 5.00
Begin VB.UserControl ucImageList 
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   690
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucImageList.ctx":0000
   PropertyPages   =   "ucImageList.ctx":117A
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   46
   ToolboxBitmap   =   "ucImageList.ctx":118D
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module      : ucImageList
' DateTime    : 14/03/08 11:40
' Author      : Cobein
' Mail        : cobein27@hotmail.com
' Purpose     : Custom Imagelist that supports many formats
' Requirements: GDI Plus
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'
' Credits     : LaVolpe, Paul Caton and http://www.activevb.de
'
' History     : 14/03/08 First Cut
'               20/03/08 Some fixes and additions in ppg
'               21/08/12 Addes file sorting
'---------------------------------------------------------------------------------------
Option Explicit

Private Const GWL_WNDPROC       As Long = -4
Private Const GW_OWNER          As Long = 4
Private Const WS_CHILD          As Long = &H40000000
Private Const UnitPixel         As Long = &H2&

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Type tFiles
    sName                       As String
    bvData()                    As Byte
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
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
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private c_lhWnd                 As Long
Private c_tvFiles()             As tFiles
Private c_bvData()              As Byte
Private hWinSafeGDI             As Long

Public Sub About()
Attribute About.VB_UserMemId = -552
    Call MsgBox("Cobein ucImageList Control, Version 0.2" & _
    vbNewLine & vbNewLine & _
    "http://www.ClassicVisualBasic.com", , "About ucImageList Control")
End Sub

'==================================================================================
'////////////////////////////          METHODS           \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Public Property Get ImageCount() As Long
    If IsArrayDim(VarPtrArray(c_tvFiles)) Then
        ImageCount = UBound(c_tvFiles)
    Else
        ImageCount = -1
    End If
End Property

Public Function GetStream(ByVal lIndex As Long) As Byte()
    If IsArrayDim(VarPtrArray(c_bvData)) Then
        If lIndex >= 0 Then
            If lIndex <= UBound(c_bvData) Then
                GetStream = c_tvFiles(lIndex).bvData
            End If
        End If
    End If
End Function

Public Function GetFileName(ByVal lIndex As Long) As String
    If IsArrayDim(VarPtrArray(c_bvData)) Then
        If lIndex >= 0 Then
            If lIndex <= UBound(c_bvData) Then
                GetFileName = c_tvFiles(lIndex).sName
            End If
        End If
    End If
End Function

Public Function SaveToFile(ByVal sFile As String, ByVal lIndex As Long) As Boolean
    Dim iFile       As Integer
    Dim bvData()    As Byte
        
    On Local Error GoTo SaveToFile_Error
        
    bvData = GetStream(lIndex)
    
    If IsArrayDim(VarPtrArray(c_bvData)) Then

        iFile = FreeFile
        Open sFile For Binary Access Write As iFile
        Put iFile, , c_bvData
        Close iFile
        SaveToFile = True
    End If
    
    Exit Function
SaveToFile_Error:
End Function

'==================================================================================
'////////////////////////////       PROPERTY PAGE        \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Friend Function ppgGetStream() As Byte()
    ppgGetStream = c_bvData
End Function

Friend Function ppgSetStream(ByRef bvData() As Byte)
    c_bvData = bvData
    Call PropertyChanged("bvData")
End Function

'==================================================================================
'////////////////////////////        USER CONTROL        \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Private Sub UserControl_Paint()
    With UserControl
        .Width = .ScaleX(38, vbPixels, vbTwips)
        .Height = .ScaleY(38, vbPixels, vbTwips)
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    c_lhWnd = UserControl.ContainerHwnd
    hWinSafeGDI = ManageGDIToken(c_lhWnd)
    
    With PropBag
        If CBool(.ReadProperty("bData", False)) Then
            c_bvData() = .ReadProperty("bvData")
            Call UnpackData(c_bvData)
        End If
    End With
End Sub

Private Sub UserControl_Terminate()
DestroyWindow hWinSafeGDI
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        If IsArrayDim(VarPtrArray(c_bvData)) Then
            Call .WriteProperty("bvData", c_bvData)
            Call .WriteProperty("bData", True)
        Else
            Call .WriteProperty("bData", False)
        End If
    End With
End Sub

'==================================================================================
'////////////////////////////      HELPER FUNCTIONS      \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress    As Long
    
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Function UnpackData(ByRef bvData() As Byte) As Boolean
    Dim cBag        As New PropertyBag
    Dim i           As Long
    Dim lCount      As Long
    
    If Not IsArrayDim(VarPtrArray(bvData)) Then
        Exit Function
    End If
    
    With cBag
        .Contents = bvData
        lCount = .ReadProperty("Index", 0)
    
        If lCount = 0 Then Exit Function
        lCount = lCount - 1
    
        ReDim c_tvFiles(lCount)
    
        For i = 0 To lCount
            c_tvFiles(i).bvData = .ReadProperty("FILE_" & i)
            c_tvFiles(i).sName = .ReadProperty("NAME_" & i)
        Next
    End With
    
    UnpackData = True
End Function

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
