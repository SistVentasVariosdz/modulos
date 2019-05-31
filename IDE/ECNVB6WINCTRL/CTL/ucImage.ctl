VERSION 5.00
Begin VB.UserControl ucImage 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   PropertyPages   =   "ucImage.ctx":0000
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   134
   ToolboxBitmap   =   "ucImage.ctx":000F
   Windowless      =   -1  'True
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   240
   End
End
Attribute VB_Name = "ucImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module      : ucImage
' DateTime    : 04/03/2008 11:00
' Author      : Cobein
' Mail        : cobein27@hotmail.com
' Purpose     : Simple Image control replacement (Beta)
' Requirements: GDI Plus
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'
' Credits     : LaVolpe, Paul Caton and http://www.activevb.de
'
' History     : 04/03/2008 Alpha realease
'               06/03/2008 Alpha 1
'               06/03/2008 Alpha 2
'               07/03/2008 Beta Release, added properties and methods
'               07/03/2008 Added bright and contrast
'               20/03/2008 Added 5 stretchig methods
'               22/03/2008 Major changes
'               01/10/2009 Fix Incremet in Memory   'Leandro Ascierto
'               25/10/2009 create shape region      'Leandro Ascierto Copy to LaVolpe
'               26/10/2009 LoadImageFromUrl         'Leandro Ascierto
'               16/11/2009 Rem zTerminate in Usercontrol_Terminate
'               23/11/2009 Remove Api SetTimer, Add ControlTimer to Prevent Crash
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0

Private Const GWL_WNDPROC       As Long = -4
Private Const GW_OWNER          As Long = 4
Private Const WS_CHILD          As Long = &H40000000
Private Const UnitPixel         As Long = &H2&

Private Const DT_VCENTER As Long = &H4
Private Const DT_CENTER As Long = &H1
Private Const DT_SINGLELINE As Long = &H20

Private Const InterpolationModeNearestNeighbor      As Long = &H5&
Private Const InterpolationModeHighQualityBicubic   As Long = &H7&
Private Const InterpolationModeHighQualityBilinear  As Long = &H6&

Private Enum ColorAdjustType
    ColorAdjustTypeDefault = 0
    ColorAdjustTypeBitmap = 1
    ColorAdjustTypeBrush = 2
    ColorAdjustTypePen = 3
    ColorAdjustTypeText = 4
    ColorAdjustTypeCount = 5
    ColorAdjustTypeAny = 6
End Enum

Private Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0
    ColorMatrixFlagsSkipGrays = 1
    ColorMatrixFlagsAltGray = 2
End Enum

Enum eScaleMode
    eActualSize
    eStretch
    eScaleDown
    eScale
    eScaleUp
End Enum

Private Type RECTF
    nLeft                       As Single
    nTop                        As Single
    nWidth                      As Single
    nHeight                     As Single
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Type POINTAPI
    X                           As Long
    Y                           As Long
End Type

Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

Private Type COLORMATRIX
    m(0 To 4, 0 To 4)           As Single
End Type

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type

Private Type SafeArray
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiPalette As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, ByRef hbmReturn As Long, ByVal Background As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
Private Declare Function GdipGetImageBounds Lib "GdiPlus.dll" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
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
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As ColorAdjustType, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal flags As ColorMatrixFlags) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As Any, ByVal un As Long, ByRef Pointer As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long


Public Event Click(ByVal Button As Integer)
Public Event DblClick(ByVal Button As Integer)
Public Event MouseExit()
Public Event MouseEnter()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event DownloadProgress(BytesMax As Long, BytesLeidos As Long)
Public Event DownloadComplete()
Public Event DownloadError()


Private c_lBtnClickTracker      As Long
Private c_lBitmap               As Long
Private c_lAttributes           As Long
Private c_lWidth                As Long
Private c_lHeight               As Long
Private c_bvData()              As Byte
Private c_sFilename             As String
Private c_eScale                As eScaleMode
Private c_bIn                   As Boolean
Private c_tPT                   As POINTAPI
Private c_lhWnd                 As Long
Private c_lContrast             As Long
Private c_lBrightness           As Long
Private c_lAlpha                As Long
Private c_bGrayScale            As Boolean
Private c_bFlipH                As Boolean
Private c_bFlipV                As Boolean
Private c_lAngle                As Long
Private m_pointer               As Long
Private m_Handle                As Long
Private m_hDC                   As Long
Private hRgn                    As Long
Private c_bUseRgn               As Boolean
Private c_AsyncProp             As AsyncProperty
Private c_bDrawProgress         As Boolean

Public Sub About()
Attribute About.VB_UserMemId = -552
    Call MsgBox("Cobein ucImage Control, Version 0.3" & _
       vbNewLine & vbNewLine & _
       "http://www.ClassicVisualBasic.com", , "About ucImage Control")
End Sub

Public Property Get UseGraficRegion() As Boolean
    UseGraficRegion = c_bUseRgn
End Property

Public Property Let UseGraficRegion(ByVal bUseRgn As Boolean)
    If hRgn <> 0 Then DeleteObject hRgn
    hRgn = 0
    c_bUseRgn = bUseRgn
    Call PropertyChanged("bUseGraficRegion")
    Call Me.Refresh
End Property

'==================================================================================
'////////////////////////////         PROPERTIES         \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get FlipHorizontal() As Boolean
    FlipHorizontal = c_bFlipH
End Property

Public Property Let FlipHorizontal(ByVal bFlipH As Boolean)
    c_bFlipH = bFlipH
    Call PropertyChanged("bFlipH")
    Call Me.Refresh
End Property

Public Property Get FlipVertical() As Boolean
    FlipVertical = c_bFlipV
End Property

Public Property Let FlipVertical(ByVal bFlipV As Boolean)
    c_bFlipV = bFlipV
    Call PropertyChanged("bFlipV")
    Call Me.Refresh
End Property

Public Property Get GrayScale() As Boolean
    GrayScale = c_bGrayScale
End Property

Public Property Let GrayScale(ByVal bGrayScale As Boolean)
    c_bGrayScale = bGrayScale
    Call PropertyChanged("bGrayScale")
    Call UserControl.Refresh
End Property

Public Property Get ScaleMode() As eScaleMode
    ScaleMode = c_eScale
End Property

Public Property Let ScaleMode(ByVal eScaleMode As eScaleMode)
    c_eScale = eScaleMode
    Call PropertyChanged("eScale")
    Call Me.Refresh
End Property

Public Property Get Brightness() As Long
    Brightness = c_lBrightness
End Property

Public Property Let Brightness(ByVal lBrightness As Long)
    c_lBrightness = lBrightness
    Call PropertyChanged("lBrightness")
    Call UserControl.Refresh
End Property

Public Property Get Contrast() As Long
    Contrast = c_lContrast
End Property

Public Property Let Contrast(ByVal lContrast As Long)
    c_lContrast = lContrast
    Call PropertyChanged("lContrast")
    Call Me.Refresh
End Property

Public Property Get Alpha() As Long
    Alpha = c_lAlpha
End Property

Public Property Let Alpha(ByVal lAlpha As Long)
    c_lAlpha = lAlpha
    Call PropertyChanged("lAlpha")
    Call Me.Refresh
End Property

Public Property Get Angle() As Long
    Angle = c_lAngle
End Property

Public Property Let Angle(ByVal lAngle As Long)
    c_lAngle = lAngle
    Call PropertyChanged("lAngle")
    Call Me.Refresh
End Property

Public Property Get PictureWidth() As Long
    PictureWidth = c_lWidth
End Property

Public Property Get PictureHeight() As Long
    PictureHeight = c_lHeight
End Property


Public Property Let Enabled(Enable As Boolean)
    UserControl.Enabled = Enable
    PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

'==================================================================================
'////////////////////////////          METHODS           \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Public Sub Refresh()
    Call UserControl.Refresh
    If c_bUseRgn Then Call CreateRgn
End Sub

Public Function PaintPicture( _
       ByVal lhDC As Long, _
       ByVal dstX As Long, _
       ByVal dstY As Long, _
       Optional ByVal dstWidth As Long, _
       Optional ByVal dstHeight As Long, _
       Optional ByVal SrcX As Long, _
       Optional ByVal SrcY As Long, _
       Optional ByVal srcWidth As Long, _
       Optional ByVal srcHeight As Long) As Boolean
       
    PaintPicture = RenderTo(lhDC, _
       dstX, dstY, dstWidth, dstHeight, _
       SrcX, SrcY, srcWidth, srcHeight)
       
End Function

Public Function IconHandle() As Long
    Call GdipCreateHICONFromBitmap(c_lBitmap, IconHandle)
End Function

Public Function GetStream() As Byte()
    GetStream = c_bvData
End Function

Public Function LoadImageFromStream(ByRef bvStream() As Byte) As Boolean
    If LoadFromStream(bvStream) Then
        c_bvData() = bvStream
        LoadImageFromStream = True
        Call Me.Refresh
    End If
End Function

Public Function SaveToFile(ByVal sFile As String) As Boolean
    Dim iFile       As Integer
    
    On Local Error GoTo SaveToFile_Error

    iFile = FreeFile
    Open sFile For Binary Access Write As iFile
    Put iFile, , c_bvData
    Close iFile
    SaveToFile = True
    
    Exit Function
SaveToFile_Error:
End Function

Public Function GetFileName() As String
    GetFileName = c_sFilename
End Function

Public Function LoadImageFromFile(ByVal sFile As String) As Boolean
    LoadImageFromFile = ppgLoadStream(sFile)
End Function

Public Function LoadImageFromRes( _
       ByVal ResIndex As Variant, _
       ByVal ResSection As Variant, _
       Optional VBglobal As IUnknown) As Boolean
    
    Dim bvData()    As Byte
    Dim oVBglobal   As VB.Global
    
    On Local Error GoTo LoadImageFromCustomRes_Error

    If VBglobal Is Nothing Then
        Set oVBglobal = VB.Global
    ElseIf TypeOf VBglobal Is VB.Global Then
        Set oVBglobal = VBglobal
    ElseIf VBglobal Is Nothing Then
        Set oVBglobal = VB.Global
    End If
    
    bvData = oVBglobal.LoadResData(ResIndex, ResSection)
    
    LoadImageFromRes = LoadFromStream(bvData)

    Call UserControl.Cls
    Call LoadFromStream(bvData)
    Call UserControl_Paint

LoadImageFromCustomRes_Error:
End Function

'==================================================================================
'////////////////////////////       PROPERTY PAGE        \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Friend Function ppgLoadStream(ByVal sFile As String) As Boolean
    Dim iFile       As Integer
    Dim bvData()    As Byte
    Dim svName()    As String
    
    On Local Error GoTo LoadStream_Error

    iFile = FreeFile
    Open sFile For Binary Access Read As iFile
    ReDim bvData(LOF(iFile) - 1)
    Get iFile, , bvData
    Close iFile
    
    svName = Split(sFile, "\")
    c_sFilename = svName(UBound(svName))
    c_bvData() = bvData
    
    Call PropertyChanged("bvData")
    Call PropertyChanged("Filename")
    
    Call LoadFromStream(bvData)
    Me.Refresh

    ppgLoadStream = True
LoadStream_Error:
End Function

Friend Function ppgGetFilename() As String
    ppgGetFilename = c_sFilename
End Function

Private Sub Timer1_Timer()
    On Error Resume Next
    If IsMouseInArea = False Then
        Timer1.Interval = 0
        c_bIn = False
        RaiseEvent MouseExit
    End If
    
End Sub

'==================================================================================
'////////////////////////////        USER CONTROL        \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_Click()
    If IsMouseInArea Then
        RaiseEvent Click(c_lBtnClickTracker \ &H10)
    End If
End Sub

Private Sub UserControl_DblClick()
    If IsMouseInArea Then
        RaiseEvent DblClick(c_lBtnClickTracker \ &H10)
    End If
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit

    If Ambient.UserMode Then
        Dim PT  As POINTAPI
        Call GetCursorPos(c_tPT)
        Call ClientToScreen(c_lhWnd, PT)
        c_tPT.X = c_tPT.X - PT.X - X
        c_tPT.Y = c_tPT.Y - PT.Y - Y
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If c_bUseRgn And hRgn <> 0 Then
        If PtInRegion(hRgn, X, Y) = 0 Then
            Exit Sub
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, X, Y)

    If Not c_bIn Then
        c_bIn = True
        RaiseEvent MouseEnter
        Timer1.Interval = 10
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If c_bUseRgn And hRgn <> 0 Then
        If PtInRegion(hRgn, X, Y) = 0 Then
            Exit Sub
        End If
    End If

    c_lBtnClickTracker = (c_lBtnClickTracker Or Button)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If c_bUseRgn And hRgn <> 0 Then
        If PtInRegion(hRgn, X, Y) = 0 Then
            Exit Sub
        End If
    End If

    c_lBtnClickTracker = (c_lBtnClickTracker And &HF)
    If (c_lBtnClickTracker And Button) = Button Then
        c_lBtnClickTracker = (c_lBtnClickTracker Or Button * &H10)
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub DrawProgress(Max As Long, Value As Long)
    Dim Rec As RECT, RecText As RECT
    Dim hBackGroundBrush As Long, hProgressBrush As Long
    Dim lPercent As Long, lWidth As Long

    If Max = 0 Then Exit Sub
    
    lPercent = (Value * 100 / Max)
    
    hBackGroundBrush = CreateSolidBrush(&HCC9999)
    hProgressBrush = CreateSolidBrush(&HFF9900)
    
    With UserControl
        SetRect Rec, .ScaleWidth / 3, (.ScaleHeight / 2) - 7, .ScaleWidth / 1.5, (.ScaleHeight / 2) + 7
        
        RecText = Rec
        
        lWidth = (Rec.Right - Rec.Left) * lPercent / 100

        Rectangle .hdc, Rec.Left - 1, Rec.Top - 1, Rec.Right + 1, Rec.Bottom + 1
        FillRect .hdc, Rec, hBackGroundBrush

        Rec.Right = Rec.Left + lWidth
        FillRect .hdc, Rec, hProgressBrush
        
        DrawText .hdc, lPercent & "%", Len(CStr(lPercent)) + 1, RecText, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
    End With
    
    DeleteObject hBackGroundBrush
    DeleteObject hProgressBrush
End Sub



Private Sub UserControl_Paint()
    Dim lW As Long, lH As Long, lT As Long, lL As Long
    
    If Not c_AsyncProp Is Nothing Then
        If c_AsyncProp.BytesMax = c_AsyncProp.BytesRead Then
            Set c_AsyncProp = Nothing
        Else
            DrawProgress c_AsyncProp.BytesMax, c_AsyncProp.BytesRead
            Exit Sub
        End If
    End If
    
    If Not c_lBitmap = 0 Then
        On Error Resume Next
        
        With UserControl
            If c_eScale = eActualSize Then
                .Height = c_lHeight * 15
                .Width = c_lWidth * 15
            End If
        
            ScalePicture c_eScale, c_lWidth, c_lHeight, _
               .Width / 15, .Height / 15, lW, lH, lL, lT
            Call RenderTo(.hdc, lL, lT, lW, lH)
            
        End With
        
    Else
        Call DrawFrame
    End If
End Sub

Private Sub UserControl_InitProperties()
    c_lAlpha = 100
End Sub

Private Sub CreateRgn()
    On Error Resume Next
    
    Dim lW As Long, lH As Long, lT As Long, lL As Long
    
    If Not c_lBitmap = 0 Then
        
        If hRgn <> 0 Then DeleteObject hRgn
 
        With UserControl
            ScalePicture c_eScale, c_lWidth, c_lHeight, .Width / 15, .Height / 15, lW, lH, lL, lT
            Call InitializeDIB(.ScaleWidth, .ScaleHeight)
            Call RenderTo(m_hDC, lL, lT, lW, lH)
            hRgn = iparseCreateShapedRegion(m_pointer, UserControl.ScaleWidth, UserControl.ScaleHeight)
            DeleteDC m_hDC
            DeleteObject m_Handle
        End With
        
    End If

End Sub

Public Function InitializeDIB(ByVal Width As Long, ByVal Height As Long) As Boolean

    ' Creates a blank (all black, all transparent) DIB of requested height & width
    
    Dim tBMPI As BITMAPINFO, tDC As Long
    
    'DestroyDIB ' clear any pre-existing dib
    
    If Width < 0& Then Exit Function
    If Height = 0& Then
        Exit Function
    ElseIf Height < 0& Then
        Height = Abs(Height) ' no top-down dibs
    End If
    
    On Error Resume Next
    With tBMPI.bmiHeader
        .biBitCount = 32
        .biHeight = Height
        .biWidth = Width
        .biPlanes = 1
        .biSize = 40&
        .biSizeImage = .biHeight * .biWidth * 4&
    End With
    
    If Err Then
        Err.Clear
        ' only possible error would be that Width*Height*4& is absolutely huge
        Exit Function
    End If
    
    tDC = GetDC(0&) ' get screen DC
    m_Handle = CreateDIBSection(tDC, tBMPI, 0&, m_pointer, 0&, 0&)
    m_hDC = CreateCompatibleDC(tDC)
    SelectObject m_hDC, m_Handle

    ReleaseDC 0&, tDC
    
    If Not m_Handle = 0& Then    ' let's hope system resources allowed DIB creation
        InitializeDIB = True
    End If

End Function

Public Function iparseCreateShapedRegion(ByVal BitsPointer As Long, ByVal Width As Long, ByVal Height As Long) As Long

    Dim rgnRects() As RECT ' array of rectangles comprising region
    Dim rectCount As Long ' number of rectangles & used to increment above array
    Dim rStart As Long ' pixel that begins a new regional rectangle
    
    Dim X As Long, Y As Long, Z As Long ' loop counters
    
    Dim bDib() As Byte  ' the DIB bit array
    Dim tSA As SafeArray ' array overlay
    Dim rtnRegion As Long ' region handle returned by this function
    Dim lScanWidth As Long ' used to size the DIB bit array
    

    

    If Width < 1& Then Exit Function
    If Height < 1& Then Exit Function
    
    On Error GoTo CleanUp
      
    lScanWidth = Width * 4& ' how many bytes per bitmap line?
    With tSA                ' prepare array overlay
        .cbElements = 1     ' byte elements
        .cDims = 2          ' two dim array
        .pvData = BitsPointer  ' data location
        .rgSABound(0).cElements = Height
        .rgSABound(1).cElements = lScanWidth
    End With
    ' overlay now
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
    

        
        ReDim rgnRects(0 To Width * 3&) ' start with an arbritray number of rectangles
        
        ' begin pixel by pixel comparisons
        For Y = Height - 1 To 0& Step -1&
            ' the alpha byte is every 4th byte
            For X = 3& To lScanWidth - 1& Step 4&

                ' test to see if next pixel is 100% transparent
                If bDib(X, Y) = 0 Then
                    If Not rStart = 0& Then ' we're currently tracking a rectangle,
                        ' so let's close it, but see if array needs to be resized
                        If rectCount + 1& = UBound(rgnRects) Then _
                            ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                         
                         ' add the rectangle to our array
                         SetRect rgnRects(rectCount + 2&), rStart \ 4, Height - Y - 1&, X \ 4 + 1&, Height - Y
                         rStart = 0&                    ' reset flag
                         rectCount = rectCount + 1&     ' keep track of nr in use
                    End If
                
                Else
                    ' non-transparent, ensure start value set
                    If rStart = 0& Then rStart = X  ' set start point
                End If
            Next X
            If Not rStart = 0& Then
                ' got to end of bitmap without hitting another transparent pixel
                ' but we're tracking so we'll close rectangle now
               
               ' see if array needs to be resized
               If rectCount + 1& = UBound(rgnRects) Then _
                   ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                   
                ' add the rectangle to our array
                SetRect rgnRects(rectCount + 2&), rStart \ 4, Height - Y - 1&, Width, Height - Y
                rStart = 0&                     ' reset flag
                rectCount = rectCount + 1&      ' keep track of nr in use
            End If
        Next Y
        


    ' remove the array overlay
    CopyMemory ByVal VarPtrArray(bDib()), 0&, 4&
        
    On Error Resume Next
    
    ' check for failure & engage backup plan if needed
    If Not rectCount = 0 Then
        ' there were rectangles identified, try to create the region in one step
        'rtnRegion = local_CreatePartialRegion(rgnRects(), 2&, rectCount + 1&, 0&, Width)
        
        With rgnRects(0) ' bytes 0-15
            .Left = 32&                     ' length of region header in bytes
            .Top = 1&                       ' required cannot be anything else
            .Right = rectCount              ' number of rectangles for the region
            .Bottom = .Right * 16&          ' byte size used by the rectangles; can be zero
        End With
        
        With rgnRects(1) ' bytes 16-31 bounding rectangle identification
            .Top = rgnRects(2).Top                     ' top
            .Right = Width                             ' right
            .Bottom = rgnRects(rectCount + 1).Bottom   ' bottom
        End With
        
        ' call function to create region from our byte (RECT) array
        rtnRegion = ExtCreateRegion(ByVal 0&, (rgnRects(0).Right + 2&) * 16&, rgnRects(0))

        
    End If

CleanUp:
    Erase rgnRects()

    If Err Then ' failure; probably low on resources
        If Not rtnRegion = 0& Then DeleteObject rtnRegion
        Err.Clear
        Debug.Print "Error Region"
    Else
        iparseCreateShapedRegion = rtnRegion
    End If


End Function


Private Sub UserControl_Resize()
    If c_bUseRgn Then Call CreateRgn
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    c_lhWnd = UserControl.ContainerHwnd
    Call ManageGDIToken(c_lhWnd)
        
    With PropBag
        c_sFilename = .ReadProperty("Filename", vbNullString)
        c_eScale = .ReadProperty("eScale", 0)
        c_lContrast = .ReadProperty("lContrast", 100)
        c_lBrightness = .ReadProperty("lBrightness", 0)
        c_lAlpha = .ReadProperty("lAlpha", 100)
        c_bGrayScale = .ReadProperty("bGrayScale", False)
        c_lAngle = .ReadProperty("lAngle", 0)
        c_bFlipH = .ReadProperty("bFlipH", False)
        c_bFlipV = .ReadProperty("bFlipV", False)
        c_bUseRgn = .ReadProperty("bUseGraficRegion", False)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        
        If CBool(.ReadProperty("bData", False)) Then
            c_bvData() = .ReadProperty("bvData")
            If c_lBitmap = 0 Then
                Dim bvData() As Byte
                bvData = c_bvData
                Call LoadFromStream(bvData)
            End If
        End If

    End With

    If c_bUseRgn Then Call CreateRgn

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        If IsArrayDim(VarPtrArray(c_bvData)) Then
            Call .WriteProperty("bvData", c_bvData)
            Call .WriteProperty("bData", True)
        Else
            Call .WriteProperty("bData", False)
        End If
        Call .WriteProperty("Filename", c_sFilename)
        Call .WriteProperty("eScale", c_eScale)
        Call .WriteProperty("lContrast", c_lContrast)
        Call .WriteProperty("lBrightness", c_lBrightness)
        Call .WriteProperty("lAlpha", c_lAlpha)
        Call .WriteProperty("bGrayScale", c_bGrayScale)
        Call .WriteProperty("lAngle", c_lAngle)
        Call .WriteProperty("bFlipH", c_bFlipH)
        Call .WriteProperty("bFlipV", c_bFlipV)
        Call .WriteProperty("bUseGraficRegion", c_bUseRgn, False)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
    End With

End Sub

Private Sub UserControl_Terminate()
    Call ClearUp
    
    'Call zTerminate

    If hRgn <> 0 Then DeleteObject hRgn
End Sub

'==================================================================================
'////////////////////////////      HELPER FUNCTIONS      \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Private Sub DrawFrame()
    Dim lhPen As Long
    On Error Resume Next
    If Not Ambient.UserMode Then
        With UserControl
            lhPen = CreatePen(2, 1, &HFF0000)
            Call SelectObject(.hdc, lhPen)
            Call Rectangle(.hdc, 0, 0, .Width / 15, .Height / 15)
            Call DeleteObject(lhPen)
        End With
    End If
End Sub

Private Function RenderTo( _
       ByVal lhDC As Long, _
       ByVal dstX As Long, _
       ByVal dstY As Long, _
       Optional ByVal dstWidth As Long, _
       Optional ByVal dstHeight As Long, _
       Optional ByVal SrcX As Long, _
       Optional ByVal SrcY As Long, _
       Optional ByVal srcWidth As Long, _
       Optional ByVal srcHeight As Long) As Boolean

    Dim hGraphics       As Long
    Dim hAttributes     As Long
    Dim bvData()        As Byte
        
    Dim dBrightness     As Double
    Dim dContrast       As Double
    Dim dAlpha          As Double
    Dim tMatrixColor    As COLORMATRIX
    Dim tMatrixGray     As COLORMATRIX
    
    bvData = c_bvData
    Call LoadFromStream(bvData)

    If c_lBitmap = 0 Then Exit Function
    
    If dstWidth = 0 Then dstWidth = c_lWidth
    If dstHeight = 0 Then dstHeight = c_lHeight
    If srcWidth = 0 Then srcWidth = c_lWidth
    If srcHeight = 0 Then srcHeight = c_lHeight
    
    dBrightness = ValidateValue(c_lBrightness, True)
    dContrast = ValidateValue(c_lContrast)
    dAlpha = ValidateValue(c_lAlpha)
    
    If GdipCreateFromHDC(lhDC, hGraphics) = 0 Then
        
        With tMatrixColor
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
            .m(4, 4) = 1
            
            If Not dContrast = 0 Then
                .m(0, 0) = 1 + dContrast
                .m(1, 1) = .m(0, 0)
                .m(2, 2) = .m(0, 0)
            End If
            
            If Not dBrightness = 0 Then
                .m(0, 4) = dBrightness
                .m(1, 4) = .m(0, 4)
                .m(2, 4) = .m(0, 4)
            End If
     
            If Not dAlpha = 100 Then
                .m(3, 3) = dAlpha
            End If
            
            If c_bGrayScale Then
                .m(0, 0) = 0.299
                .m(1, 0) = .m(0, 0)
                .m(2, 0) = .m(0, 0)
                .m(0, 1) = 0.587
                .m(1, 1) = .m(0, 1)
                .m(2, 1) = .m(0, 1)
                .m(0, 2) = 0.114
                .m(1, 2) = .m(0, 2)
                .m(2, 2) = .m(0, 2)
            End If
        End With

        If c_bFlipH Then Call GdipImageRotateFlip(c_lBitmap, 4&)
        If c_bFlipV Then Call GdipImageRotateFlip(c_lBitmap, 6&)
                            
        If GdipCreateImageAttributes(hAttributes) = 0 Then
                
            If GdipSetImageAttributesColorMatrix( _
               hAttributes, ColorAdjustTypeDefault, True, _
               tMatrixColor, tMatrixGray, _
               ColorMatrixFlagsDefault) = 0 Then
           
                If c_lAngle = 0 Then
                    If GdipDrawImageRectRectI( _
                       hGraphics, _
                       c_lBitmap, _
                       dstX, dstY, dstWidth, dstHeight, _
                       SrcX, SrcY, srcWidth, srcHeight, _
                       UnitPixel, _
                       hAttributes) = 0 Then
                        RenderTo = True
                    End If
                Else
                    If GdipRotateWorldTransform(hGraphics, c_lAngle + 180, 0) = 0 Then
                        Call GdipTranslateWorldTransform( _
                           hGraphics, _
                           dstX + (dstWidth \ 2), dstY + (dstHeight \ 2), _
                           1)
                    End If
                    If GdipDrawImageRectRectI( _
                       hGraphics, _
                       c_lBitmap, _
                       dstWidth \ 2, dstHeight \ 2, -dstWidth, -dstHeight, _
                       SrcX, SrcY, srcWidth, srcHeight, _
                       UnitPixel, _
                       hAttributes) = 0 Then
                        RenderTo = True
                    End If
                End If
            End If
                
            Call GdipDisposeImageAttributes(hAttributes)
        End If
        
        Call GdipDeleteGraphics(hGraphics)
    End If
    
End Function

Private Function ValidateValue(ByVal dVal As Double, Optional bNetative As Boolean) As Double
    If dVal < 0 And bNetative = False Then
        ValidateValue = 0
        Exit Function
    ElseIf dVal > 100 Then
        dVal = 100
    End If
    ValidateValue = dVal / 100
End Function

Public Function LoadImageFromURL(ByVal sUrl As String, Optional ByVal UseCache As Boolean, Optional ByVal DrawProgress As Boolean = True) As Boolean
    On Error GoTo PropErr
    
    c_bDrawProgress = DrawProgress
    
    If Left(LCase(sUrl), 7) <> "http://" Then
        sUrl = "http://" & sUrl
    End If
    
    If UseCache = False Then
        Call AsyncRead(sUrl, vbAsyncTypeByteArray, sUrl, vbAsyncReadForceUpdate)
    Else
        Call AsyncRead(sUrl, vbAsyncTypeByteArray, sUrl)
    End If
    
    LoadImageFromURL = True
    
    Exit Function

PropErr:

End Function

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    On Error Resume Next
    If c_bDrawProgress Then
        If c_AsyncProp Is Nothing Then Set c_AsyncProp = AsyncProp
        UserControl.Refresh
    End If
    RaiseEvent DownloadProgress(AsyncProp.BytesMax, AsyncProp.BytesRead)
End Sub



Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error GoTo PropErr
    
    If LoadImageFromStream(AsyncProp.Value) Then
        RaiseEvent DownloadComplete
    Else
        RaiseEvent DownloadError
        
    End If
    
    Set c_AsyncProp = Nothing
    
    Exit Sub
PropErr:
    RaiseEvent DownloadError
End Sub



Private Function LoadFromStream(ByRef bvData() As Byte) As Boolean
    Dim IStream     As IUnknown
    Dim lhBitmap    As Long
    Dim TR          As RECTF
    
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function

    Call ClearUp
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, c_lBitmap) = 0 Then
            LoadFromStream = True
            Call GdipGetImageBounds(c_lBitmap, TR, UnitPixel)
            c_lWidth = TR.nWidth
            c_lHeight = TR.nHeight
        End If
    End If

    Set IStream = Nothing
End Function

Private Sub ClearUp()
    If Not c_lBitmap = 0 Then
        Call GdipDisposeImage(c_lBitmap)
        c_lBitmap = 0: c_lWidth = 0: c_lHeight = 0
    End If
End Sub

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Function ScalePicture( _
       ByVal eScaleMode As eScaleMode, _
       ByVal lSrcWidth As Long, _
       ByVal lSrcHeight As Long, _
       ByVal lDstWidth As Long, _
       ByVal lDstHeight As Long, _
       ByRef lNewWidth As Long, _
       ByRef lNewHeight As Long, _
       ByRef lNewLeft As Long, _
       ByRef lNewTop As Long)

    Dim dHRatio As Double
    Dim dVRatio As Double
    Dim dRatio  As Double
    
    dHRatio = lSrcWidth / lDstWidth
    dVRatio = lSrcHeight / lDstHeight
     
    Select Case eScaleMode
        Case eActualSize
            lNewWidth = lSrcWidth
            lNewHeight = lSrcHeight
        Case eStretch
            lNewWidth = lDstWidth
            lNewHeight = lDstHeight
        Case eScaleDown
            If dHRatio > 1 Or dVRatio > 1 Then
                If dHRatio > dVRatio Then
                    dRatio = dHRatio
                Else
                    dRatio = dVRatio
                End If
            Else
                lNewWidth = lSrcWidth
                lNewHeight = lSrcHeight
            End If
        Case eScale
            If dHRatio > dVRatio Then
                dRatio = dHRatio
            Else
                dRatio = dVRatio
            End If
        Case eScaleUp
            If dHRatio < dVRatio Then
                dRatio = dHRatio
            Else
                dRatio = dVRatio
            End If
    End Select
    
    If Not dRatio = 0 Then
        lNewWidth = lSrcWidth / dRatio
        lNewHeight = lSrcHeight / dRatio
    End If
    
    lNewLeft = (lDstWidth - lNewWidth) / 2
    lNewTop = (lDstHeight - lNewHeight) / 2
End Function

Public Function IsMouseInArea() As Boolean
    Dim PT As POINTAPI
    Dim CPT As POINTAPI
    Dim TR As RECT
    Dim bArea As Boolean
    
    Call GetCursorPos(PT)
    Call ClientToScreen(c_lhWnd, CPT)
    
    CPT.X = PT.X - CPT.X - c_tPT.X
    CPT.Y = PT.Y - CPT.Y - c_tPT.Y

    If c_bUseRgn And hRgn <> 0 Then
        bArea = PtInRegion(hRgn, CPT.X, CPT.Y)
    Else
        Call SetRect(TR, 0, 0, UserControl.Width / 15, UserControl.Height / 15)
        bArea = PtInRect(TR, CPT.X, CPT.Y)
    End If
    
    If bArea And WindowFromPoint(PT.X, PT.Y) = c_lhWnd Then
        IsMouseInArea = True
    End If

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

