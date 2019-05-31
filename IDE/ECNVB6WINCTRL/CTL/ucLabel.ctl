VERSION 5.00
Begin VB.UserControl ucLabel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1650
   DrawStyle       =   2  'Dot
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   110
End
Attribute VB_Name = "ucLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CopyRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSourceRect As RECT) As Long

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Const COLOR_BTNFACE As Long = 15
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_BTNTEXT As Long = 18

Private Const BDR_INNER As Long = &HC
Private Const BDR_OUTER As Long = &H3
Private Const BDR_RAISED As Long = &H5
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_SUNKEN As Long = &HA
Private Const BDR_SUNKENINNER As Long = &H8
Private Const BDR_SUNKENOUTER As Long = &H2

Private Const BF_ADJUST As Long = &H2000
Private Const BF_BOTTOM As Long = &H8
Private Const BF_DIAGONAL As Long = &H10
Private Const BF_FLAT As Long = &H4000
Private Const BF_LEFT As Long = &H1
Private Const BF_MIDDLE As Long = &H800
Private Const BF_MONO As Long = &H8000
Private Const BF_RIGHT As Long = &H4
Private Const BF_SOFT As Long = &H1000
Private Const BF_TOP As Long = &H2
Private Const BF_TOPLEFT As Long = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT As Long = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT As Long = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT As Long = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDBOTTOMLEFT As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT As Long = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDTOPRIGHT As Long = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const EDGE_BUMP As Long = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED As Long = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_MULTILINE As Long = (&H1)
Private Const DT_RIGHT As Long = &H2
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_TOP As Long = &H0
Private Const DT_VCENTER As Long = &H4
Private Const DT_WORDBREAK As Long = &H10

Public Enum AlignStyle
    [Left Justified] = 0
    [Right Justified] = 1
    [Center Justified] = 2
End Enum

'Public Enum TipoBackColor
'    [Opaco] = 0
'    [Transparente] = 1
'End Enum

'Dim m_TipoFondo As TipoBackColor
Dim m_Caption As String
Dim m_Alignment As AlignStyle
Dim m_Autosize As Boolean
Dim m_WordWrap As Boolean
Dim m_Enabled As Boolean
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_ShadowColor As OLE_COLOR
Dim m_Font As Font

Event Click()
Attribute Click.VB_Description = "Trigger when the label is clicked."
Attribute Click.VB_MemberFlags = "200"
Event DblClick()
Attribute DblClick.VB_Description = "Triggered when the label is double clicked."
Event Change()
Attribute Change.VB_Description = "Triggered when the caption has been changed."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Trigger when the mouse button is pushed."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Triggered when the mouse is moved over the label."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Triggered when the mouse is released."




Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/Sets whether or not the label is allowed to be clicked."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    DrawLabel
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/Sets the font used for the caption."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Font = m_Font
End Property

Public Property Set Font(New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
    DrawLabel
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/Sets whether the text is automatically wrapped to the next line or not."
Attribute WordWrap.VB_ProcData.VB_Invoke_Property = ";Behavior"
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(New_WordWrap As Boolean)
    m_WordWrap = New_WordWrap
    PropertyChanged "WordWrap"
    DrawLabel
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/Sets the text color."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    DrawLabel
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/Set the color used in the background of the label."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    DrawLabel
End Property

Public Property Get ShadowColor() As OLE_COLOR
Attribute ShadowColor.VB_Description = "Returns/Sets the color of the shadow."
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(New_ShadowColor As OLE_COLOR)
    m_ShadowColor = New_ShadowColor
    PropertyChanged "ShadowColor"
    DrawLabel
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/Sets what text is displayed in the label."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_MemberFlags = "200"
    Caption = m_Caption
End Property

Public Property Let Caption(New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    RaiseEvent Change
    DrawLabel
End Property

Public Property Get Alignment() As AlignStyle
Attribute Alignment.VB_Description = "Returns/Sets the font used for the caption."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(New_Alignment As AlignStyle)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    DrawLabel
End Property

'Public Property Get BackStyle() As TipoBackColor
'    BackStyle = m_TipoFondo
'End Property
'
'Public Property Let BackStyle(New_TipoFondo As TipoBackColor)
'    m_TipoFondo = New_TipoFondo
'    PropertyChanged "BackStyle"
'    DrawLabel
'End Property

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "If set to True the control will automatically resize itself to fit the caption."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoSize = m_Autosize
End Property

Public Property Let AutoSize(New_Autosize As Boolean)
    m_Autosize = New_Autosize
    PropertyChanged "Autosize"
    DrawLabel
End Property

Private Sub DrawLabel()
    On Error Resume Next
    
    Dim hRect As RECT
    Dim hFlags As Long, tFlag As Long, jFlag As Long
    UserControl.Cls
    SetRect hRect, 3, 3, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3

    Select Case m_Alignment
        Case 0: jFlag = DT_LEFT
        Case 1: jFlag = DT_RIGHT
        Case 2: jFlag = DT_CENTER
    End Select
'
'    Select Case m_TipoFondo
'    Case TipoBackColor.Opaco: Me.BackStyle = Opaco
'    Case TipoBackColor.Transparente: Me.BackStyle = Transparente
'    End Select

    Select Case m_WordWrap
        Case True: hFlags = jFlag Or DT_WORDBREAK
        Case False: hFlags = jFlag Or DT_SINGLELINE
    End Select

    If m_Autosize = True Then
        tFlag = jFlag Or DT_SINGLELINE Or DT_CALCRECT
        DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, tFlag
        UserControl.Height = hRect.bottom * Screen.TwipsPerPixelY
        UserControl.Width = hRect.right * Screen.TwipsPerPixelX
    End If

    UserControl.BackColor = GetColor(m_BackColor)
    Set UserControl.Font = m_Font

    If m_Enabled = True Then
        SetTextColor UserControl.hdc, BlendColors(GetColor(m_ShadowColor), GetColor(m_BackColor), 80)
    Else
        SetTextColor UserControl.hdc, BlendColors(GetColor(vbGrayText), GetColor(m_BackColor), 80)
    End If
    
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, -1, 0
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, -1, 0
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, -1, 0
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, 3, -1
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, 0, -1
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, 0, -1
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags


    If m_Enabled = True Then
        SetTextColor UserControl.hdc, BlendColors(GetColor(m_ShadowColor), GetColor(m_BackColor), 60)
    Else
        SetTextColor UserControl.hdc, BlendColors(GetColor(vbGrayText), GetColor(m_BackColor), 60)
    End If
    
    OffsetRect hRect, -1, 2
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, -1, 0
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, -1, 0
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, 2, -1
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, 0, -1
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags

    If m_Enabled = True Then
        SetTextColor UserControl.hdc, BlendColors(GetColor(m_ShadowColor), GetColor(m_BackColor), 40)
    Else
        SetTextColor UserControl.hdc, BlendColors(GetColor(vbGrayText), GetColor(m_BackColor), 40)
    End If
    
    OffsetRect hRect, -1, 1
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, -1, 0
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    OffsetRect hRect, 1, -1
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
    
   If m_Enabled = True Then
        SetTextColor UserControl.hdc, GetColor(m_ForeColor)
    Else
        SetTextColor UserControl.hdc, GetColor(vbGrayText)
    End If
        
   
    '--------------
    OffsetRect hRect, -1, 0
    DrawText UserControl.hdc, m_Caption, Len(m_Caption), hRect, hFlags
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
    Set m_Font = Ambient.Font
    m_Caption = Ambient.DisplayName
    m_Alignment = 0
    m_Autosize = False
    m_ForeColor = vbButtonText
    m_BackColor = vbButtonFace
    m_ShadowColor = vbButtonShadow
    m_WordWrap = True
    m_Enabled = True
    DrawLabel
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
    DrawLabel
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_Alignment = .ReadProperty("Alignment", 0)
        m_Autosize = .ReadProperty("Autosize", False)
        m_ForeColor = .ReadProperty("ForeColor", vbButtonText)
        m_BackColor = .ReadProperty("BackColor", vbButtonFace)
        m_ShadowColor = .ReadProperty("ShadowColor", vbButtonShadow)
        m_WordWrap = .ReadProperty("WordWrap", True)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        m_Enabled = .ReadProperty("Enabled", True)
    End With
End Sub

Private Sub UserControl_Show()
    DrawLabel
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
        Call .WriteProperty("Alignment", m_Alignment, 0)
        Call .WriteProperty("Autosize", m_Autosize, False)
        Call .WriteProperty("ForeColor", m_ForeColor, vbButtonText)
        Call .WriteProperty("BackColor", m_BackColor, vbButtonFace)
        Call .WriteProperty("ShadowColor", m_ShadowColor, vbButtonShadow)
        Call .WriteProperty("WordWrap", m_WordWrap, True)
        Call .WriteProperty("Font", m_Font, Ambient.Font)
        Call .WriteProperty("Enabled", m_Enabled, True)
    End With
End Sub

Private Sub UserControl_Resize()
    DrawLabel
End Sub

Private Function BlendColors(Color1 As Long, Color2 As Long, Percentage As Integer) As Long
    On Error Resume Next
    Dim R(1) As Byte, G(1) As Byte, B(1) As Byte
    Dim iR(1) As Integer, iG(1) As Integer, iB(1) As Integer
    Dim fRed As Integer, fGreen As Integer, fBlue As Integer
    Dim TempValue(1) As String
    Dim fPercentage(2) As Double
    If Percentage < 0 Then Percentage = 0
    If Percentage > 100 Then Percentage = 100
    TempValue(0) = Hex(Color1)
    If Len(TempValue(0)) < 6 Then TempValue(0) = String(6 - Len(TempValue(0)), "0") & TempValue(0)
    TempValue(1) = Hex(Color2)
    If Len(TempValue(1)) < 6 Then TempValue(1) = String(6 - Len(TempValue(1)), "0") & TempValue(1)
    R(0) = CByte("&H" & right$(TempValue(0), 2))
    G(0) = CByte("&H" & Mid$(TempValue(0), 3, 2))
    B(0) = CByte("&H" & left$(TempValue(0), 2))
    R(1) = CByte("&H" & right$(TempValue(1), 2))
    G(1) = CByte("&H" & Mid$(TempValue(1), 3, 2))
    B(1) = CByte("&H" & left$(TempValue(1), 2))
    If R(0) > R(1) Then
        iR(0) = -1
        iR(1) = R(0) - R(1)
    Else
        iR(0) = 1
        iR(1) = R(1) - R(0)
    End If
    fPercentage(0) = (iR(1) / 100) * (Percentage * iR(0))
    If G(0) > G(1) Then
        iG(0) = -1
        iG(1) = G(0) - G(1)
    Else
        iG(0) = 1
        iG(1) = G(1) - G(0)
    End If
    fPercentage(1) = (iG(1) / 100) * (Percentage * iG(0))
    If B(0) > B(1) Then
        iB(0) = -1
        iB(1) = B(0) - B(1)
    Else
        iB(0) = 1
        iB(1) = B(1) - B(0)
    End If
    fPercentage(2) = (iB(1) / 100) * (Percentage * iB(0))
    fRed = R(0) + fPercentage(0)
    fGreen = G(0) + fPercentage(1)
    fBlue = B(0) + fPercentage(2)
    BlendColors = RGB(fRed, fGreen, fBlue)
End Function

Private Function GetColor(Color As Long) As Long
    Call OleTranslateColor(Color, 0, GetColor)
End Function
