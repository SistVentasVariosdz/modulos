Attribute VB_Name = "ECNLIB03_WINEVE_RICHTXT_RESALTAR_WORD"
Option Explicit

' Win32 API Declarations
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

' General declarations
Private Const LF_FACESIZE = 32
Private Const GWL_STYLE = (-16)
Private Const SCF_SELECTION = &H1&

' Font Back Color
Private Const CFM_BACKCOLOR = &H4000000
Private Const CFE_AUTOBACKCOLOR = CFM_BACKCOLOR

' CharFormat structure, passed with SendMessage to the
' control
Private Type CHARFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To LF_FACESIZE - 1) As Byte
    wPad2 As Integer
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lLCID As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte
    bRevAuthor As Byte
    bReserved1 As Byte
End Type

' Window messages
Private Const WM_USER = &H400

' Edit messages
Private Const EM_SETCHARFORMAT = (WM_USER + 68)
Private Const EM_SETBKGNDCOLOR = (WM_USER + 67)
Private Const EM_GETCHARFORMAT = (WM_USER + 58)
Private Const WM_SETTEXT = &HC

Public Function GetSelBackColor(ByVal RTFhwnd As Long) As OLE_COLOR
    Dim udtChar As CHARFORMAT2

    ' Set BackColor mask
    udtChar.dwMask = CFM_BACKCOLOR
    udtChar.cbSize = Len(udtChar)
    ' Get character format structure from selection
    SendMessage RTFhwnd, EM_GETCHARFORMAT, SCF_SELECTION, udtChar
    ' Return the BackColour of the selection

    GetSelBackColor = udtChar.crBackColor
End Function

Public Function SetSelBackColor(ByVal RTFhwnd As Long, ByVal NewSelFontBackColor _
                                                                As OLE_COLOR)
    Dim udtChar As CHARFORMAT2
    
    ' Set the mask
    udtChar.dwMask = CFM_BACKCOLOR

    ' If the new backcolour is set to -1 then we set the
    ' RichTextbox backcolour to be "auto"
    If NewSelFontBackColor = -1 Then
        udtChar.dwEffects = CFE_AUTOBACKCOLOR
        udtChar.crBackColor = -1
    Else
        ' Set the BackColour to the new colour
        udtChar.crBackColor = TranslateColor(NewSelFontBackColor)
    End If
    
    ' We need to pass the size of the structure as a
    ' part of the structure.
    udtChar.cbSize = Len(udtChar)
    
    ' Send the SET message and the new character format
    ' structure to the RichTextbox
    SendMessage RTFhwnd, EM_SETCHARFORMAT, SCF_SELECTION, udtChar
End Function

' Used by the SetSelBackColor function, no need for the
' end user to touch this function
Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

' Pass this function a Colour in the form R*B*G and it will
' return the individual R G B components, useful for translating
' the return value of a common dialog colour picker or the
' GetSelBackColor function
Public Sub GetRGB(ByVal Colour, ByRef R, ByRef G, ByRef B)
    R = Colour Mod 256
    Colour = Colour \ 256
    G = Colour Mod 256
    Colour = Colour \ 256
    B = Colour Mod 256
End Sub


