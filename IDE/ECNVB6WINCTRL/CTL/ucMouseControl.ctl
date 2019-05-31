VERSION 5.00
Begin VB.UserControl ucMouseControl 
   CanGetFocus     =   0   'False
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   735
   ScaleWidth      =   5145
   ToolboxBitmap   =   "ucMouseControl.ctx":0000
   Begin VB.Label lblAuthor 
      Caption         =   "Advanced Mouse Control Created By Robert Engelhardt"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image imgAvatar 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   120
      Picture         =   "ucMouseControl.ctx":0312
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "ucMouseControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Rectangel usage for API
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

'Point Decliration
Private Type POINT
    X As Long
    Y As Long
End Type

'Point decliration for specific API
Private Type POINTAPI
    X As Long
    Y As Long
End Type

'API calls en mass
'cursor capturing
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

'double click speed
Private Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long

'mouse button swaping
Private Declare Function SwapMouseButton& Lib "user32" (ByVal bSwap As Long)

'cursor position
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

'visibility of cursor
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'Miscemanious data gathering
Private Const SM_CXCURSOR = 13 'Width of standard cursor
Private Const SM_CYCURSOR = 14 'Height of standard cursor
Private Const SM_MOUSEPRESENT = 19 'True is a mouse is present
Private Const SM_SWAPBUTTON = 23 'True if left and right buttons are swapped.
Private Const SM_CXDOUBLECLK = 36 'double click width
Private Const SM_CYDOUBLECLK = 37 'double click height
Private Const SM_CMOUSEBUTTONS = 43 'Number of mouse buttons.
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'Subs ----------------------------------------------------------------------------
Private Sub UserControl_Resize() 'call when re-szie is made
   imgAvatar.Move 0, 0 'move image to top left corner (not seen at runtime)
   Width = imgAvatar.Width 'don't get to wide
   Height = imgAvatar.Height 'don't get to tall
End Sub

'Functions ----------------------------------------------------------------------------
Function ReleaseCaptive()
    ClipCursor ByVal 0& 'Releases the cursor limiting by passing a null value
End Function

Public Function ShowMouse(state As Boolean)
   ShowCursor state 'show cursor with desired state
End Function

Public Function BlockAllInput(state As Boolean)
   BlockInput state 'block or gain user input
End Function
Public Function SetCustomCursor(pic As Picture)
   Screen.MouseIcon = pic 'set the custom cursor picture to be that of passed picture
End Function

Public Function CursorWidth() As Long
   CursorWidth = GetSystemMetrics(SM_CXCURSOR) 'return cursor width
End Function
Public Function CursorHeight() As Long
   CursorHeight = GetSystemMetrics(SM_CYCURSOR) 'return cursor height
End Function
Public Function DoubleClickWidth() As Long
   DoubleClickWidth = GetSystemMetrics(SM_CXDOUBLECLK) 'return width of double click box
End Function
Public Function DoubleClickHeight() As Long
   DoubleClickHeight = GetSystemMetrics(SM_CYDOUBLECLK) 'return height of double click box
End Function

Public Function ButtonCount() As Long
   ButtonCount = GetSystemMetrics(SM_CMOUSEBUTTONS) 'returnt the amount of buttons on the mouse
End Function

Public Function MouseExists() As Boolean
   MouseExists = IIf(GetSystemMetrics(SM_MOUSEPRESENT) = 1, True, False) ' return boolean value reflecting that of mouse existance
End Function
Function CaptureForm(ByVal hWnd As Long)
    
   Dim client As RECT 'client rectangle
   Dim upperleft As POINT 'uper left hand point of window
   
   GetClientRect hWnd, client 'set client rectangel to be that of passed hWnd
   upperleft.X = client.left ' set left
   upperleft.Y = client.top 'set top
    
   ClientToScreen hWnd, upperleft 'offset to screen positino of window
    
   OffsetRect client, upperleft.X, upperleft.Y 'move to match
    
   ClipCursor client 'clip the cursor to the window
    
End Function

Function CaptureRec(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Optional mode As ScaleModeConstants = vbPixels)
   
   Dim client As RECT 'client rectangel
   
   If mode <> vbPixels Then 'if need to change scale form pixels
      'set horizontal
      X1 = ScaleX(X1, mode, vbPixels)
      X2 = ScaleX(X2, mode, vbPixels)
      'set vertical
      Y1 = ScaleY(Y1, mode, vbPixels)
      Y2 = ScaleY(Y2, mode, vbPixels)
   End If
   
   SetRect client, X1, Y1, X2, Y2 'set client rectangel
   
   ClipCursor client 'clip cursor to the client rectangel

End Function

'Properties ----------------------------------------------------------------------------
Public Property Get DoubleClickSpeed() As Long
   DoubleClickSpeed = GetDoubleClickTime 'return the double click speed
End Property

Public Property Let DoubleClickSpeed(ByVal NewValue As Long)
   SetDoubleClickTime NewValue 'set the double click speed
End Property

Public Property Get SwapButtons() As Boolean
   SwapButtons = GetSystemMetrics(SM_SWAPBUTTON) 'return value stating if values are swaped
End Property

Public Property Let SwapButtons(ByVal NewValue As Boolean)
   SwapMouseButton& NewValue 'swap values if called with the new value
End Property

Public Property Get mouseX() As Long
   Dim pos As POINTAPI
   GetCursorPos pos 'get position
   mouseX = pos.X 'return x
End Property

Public Property Let mouseX(ByVal NewValue As Long)
   Dim pos As POINTAPI
   GetCursorPos pos 'get position
   SetCursorPos NewValue, pos.Y 'set new x value , old y
End Property

Public Property Get mouseY() As Long
   Dim pos As POINTAPI
   GetCursorPos pos 'get position
   mouseY = pos.Y 'return y
End Property

Public Property Let mouseY(ByVal NewValue As Long)
   Dim pos As POINTAPI
   GetCursorPos pos 'get position
   SetCursorPos pos.X, NewValue 'set old x , new y value
End Property

Public Property Get Cursor() As MousePointerConstants
   Cursor = Screen.MousePointer 'return cursor
End Property

Public Property Let Cursor(ByVal NewValue As MousePointerConstants)
   Screen.MousePointer = NewValue 'set cursor
End Property




