Attribute VB_Name = "modSistema"
Declare Function GetDeviceCaps _
        Lib "gdi32" (ByVal hdc As Long, _
                     ByVal nIndex As Long) As Long

Function fVerResol() As String

    'Funcion que devuelve la resolucion Actual
    Dim lBits As Long, lWidth As Long, lHeight As Long

    Dim hdc   As Long

    lBits = GetDeviceCaps(hdc, BITSPIXEL)
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    fVerResol = LTrim(Str(lWidth)) + "x" + Trim(Str(lHeight))
End Function

