Attribute VB_Name = "ECNLIB03_WINEVE_GRIDSCROLL"
Option Explicit

Public G_SW_UtilizaScroll As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DECLARACIONES API
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long


' YA ESTA DECLARADO EN OTRO MODULO POR ESO LO COMENTO
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CONSTANTES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_VSCROLL As Integer = &H115

Dim PrevProc As Long

' instala el hook para el control indicado
Public Sub Grilla_IniciarScroll(ElControl As Object)
    PrevProc = SetWindowLong(ElControl.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

' Remueve el Hook para el control indicado
Public Sub Grilla_DetenerScroll(ElControl As Object)
    SetWindowLong ElControl.hwnd, GWL_WNDPROC, PrevProc
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PROCEDIMIENTO PARA PROCESAR LOS MENSAJES DE WINDOWS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
       
    Dim HScroll As Long
        
        
    ' Obtiene el Hwnd de la barra de Scroll vertical del DataGrid
    HScroll = FindWindowEx(hwnd, 0, "ScrollBar", "DataGridSplitVScroll")
    
    If clase(hwnd) = "DataGridWndClass" And HScroll = 0 Then
         WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
        Exit Function
    End If
    
    If uMsg = WM_MOUSEWHEEL Then
       If G_SW_UtilizaScroll = True Then
            If clase(hwnd) = "DataGridWndClass" And HScroll <> 0 Then
                
                If wParam < 0 Then
                    ' Scroll hacia abajo
                    SendMessage hwnd, WM_VSCROLL, 1, ByVal HScroll
                Else
                    ' Mueve el scroll hacia arriba
                    SendMessage hwnd, WM_VSCROLL, 0, ByVal HScroll
                End If
            Else
                If wParam < 0 Then
                    ' Scroll hacia abajo
                    SendMessage hwnd, WM_VSCROLL, 1, ByVal 0
                Else
                    ' Mueve el scroll hacia arriba
                    SendMessage hwnd, WM_VSCROLL, 0, ByVal 0
                End If
            End If
        End If
    End If

    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
 End Function

Private Function clase(handle As Long) As String
    Dim buffer As String * 256
    Dim ret As Long
    ret = GetClassName(handle, buffer, 256)
  
    clase = Left(buffer, ret)
End Function





