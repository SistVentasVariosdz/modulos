VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB01_FUNSUB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'************************************************************************************************************************************************************
' AUTOR       : EDGAR I. CARDENAS NAVARRO
' FECHA       : 09/05/2013
' NOMBRE      : ECNLIB01
' DESCRIPCION : CLASE CONTENEDORA DE LOS PROCEDIMIENTOS Y FUNCIONES GENERALES
'************************************************************************************************************************************************************
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA FUNCION : BuscarEquiposDeRed (Common Dialog de Windows)
'--+------------------------------------------------------------------------------------------------------------------------+--

Private Const ERROR_SUCCESS As Long = 0
Private Const MAX_PATH As Long = 260
Private Const CSIDL_NETWORK As Long = &H12
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
  
'Estructura BROWSEINFO necesaria para el Api SHBrowseForFolder
Private Type BROWSEINFO 'BI
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
  
'Declaramos la función Api SHBrowseForForlder
Private Declare Function SHBrowseForFolder _
                Lib "shell32.dll" _
                Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
  
'Declaramos la función Api SHGetSpecialFolderLocation
  
Private Declare Function SHGetSpecialFolderLocation _
                Lib "shell32.dll" (ByVal hwndOwner As Long, _
                                   ByVal nFolder As Long, _
                                   pidl As Long) As Long
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA FUNCION : HABILITAR/DESHABILITAR BOTON CERRAR DE UN FORM
'--+------------------------------------------------------------------------------------------------------------------------+--
Private Declare Function GetSystemMenu _
                Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemCount _
                Lib "user32" (ByVal hMenu As Long) As Long

Private Declare Function RemoveMenu _
                Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Declare Function DrawMenuBar _
                Lib "user32" (ByVal hWnd As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private Const MF_DISABLED = &H2&
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA FUNCION : ComputerName
'--+------------------------------------------------------------------------------------------------------------------------+--
Private Declare Function GetcomputerName _
        Lib "kernel32" _
        Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA FUNCION : UserName
'--+------------------------------------------------------------------------------------------------------------------------+--
Private Declare Function GetUserName _
        Lib "advapi32.dll" _
        Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA PROCEDIMIENTO : AlinearTxTdeBoton
'--+------------------------------------------------------------------------------------------------------------------------+--
Public Enum eruAlineacion
    eruCENTRO = 1
    
    eruARRIBA = 2
    eruARRIBA_IZQUIERDA = 3
    eruARRIBA_DERECHA = 4
    
    eruABAJO = 5
    eruABAJO_IZQUIERDA = 6
    eruABAJO_DERECHA = 7
    
    eruIZQUIERDA = 8
    eruDERECHA = 9
End Enum

Private Const A_CENTER = &H300&
Private Const A_TOP = &H400&
Private Const A_TOPLEFT = &H500&
Private Const A_TOPRIGHT = &H600&
Private Const A_BOTTOM = &H800&
Private Const A_BOTTOMLEFT = &H900&
Private Const A_BOTTOMRIGHT = &HA00&
Private Const A_LEFT = &H100&
Private Const A_RIGHT = &H200&

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000

Private Declare Function GetWindowLong _
        Lib "user32" _
        Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong _
        Lib "user32" _
        Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA FUNCION : DevuelveResolucionDePantalla
'--+------------------------------------------------------------------------------------------------------------------------+--
Private Const EPOCH = "1930"   'Considera la centuria : 1970 HASTA 2070, Variar de acuerdo a la epoca del sistema
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const BITSPIXEL = 12
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA PROCEDIMIENTO : CambiaComboListaAncho
'--+------------------------------------------------------------------------------------------------------------------------+--
Private Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Const CB_SETDROPPEDWIDTH = &H160
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA PROCEDIMIENTO : CambiarComboListaLargo
'--+------------------------------------------------------------------------------------------------------------------------+--
Private Declare Function MoveWindow _
        Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA EL PROCEMIENTO : CambiarComboHeight
'--+------------------------------------------------------------------------------------------------------------------------+--
Private Const CB_SETITEMHEIGHT As Long = &H153
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA PROCEDIMIENTO : DesplegarCb
'--+------------------------------------------------------------------------------------------------------------------------+--
Private Const CB_sDespliegaCb = &H14F
'--+------------------------------------------------------------------------------------------------------------------------+--
' PARA PROCEDIMIENTO : DeGradarColoRPictuRe
'--+------------------------------------------------------------------------------------------------------------------------+--
Public Enum EruDegradaDireccion
    eRuDeGrada_HORIZONTAL = 1
    eRuDeGrada_VERTICAL = 2
End Enum

'***************************************************************************************************************************************************************************************************************************************
' Cript2 : Encripta una cadena
'***************************************************************************************************************************************************************************************************************************************
Function Cript2(ByVal sCadena As String) As String
    On Error GoTo SALTO_ERROR

    Dim bLong As Byte
    Dim bPos As Byte
    Dim sCar As String
    Dim WPASS As String

    bLong = Len(sCadena)
    For bPos = 1 To bLong
        If bPos Mod 2 <> 0 Then
            sCar = Chr(Asc(Mid(sCadena, bPos, 1)) + 3)
        Else
            sCar = Chr(Asc(Mid(sCadena, bPos, 1)) - 3)
        End If
        WPASS = WPASS + sCar
    Next bPos
    Cript2 = WPASS
    Exit Function

SALTO_ERROR:
    Cript2 = Empty
    MsgBox Err.Description, vbCritical, "ECNLIB01 - Cript2"
End Function

'***************************************************************************************************************************************************************************************************************************************
' Cript2 : Desencripta una cadena
'***************************************************************************************************************************************************************************************************************************************
Function DCript2(ByVal sCadena As String) As String
    On Error GoTo SALTO_ERROR

    Dim bLong As Byte
    Dim bPos As Byte
    Dim sCar As String
    Dim WPASS As String

    bLong = Len(sCadena)
    For bPos = 1 To bLong
        If bPos Mod 2 <> 0 Then
            sCar = Chr(Asc(Mid(sCadena, bPos, 1)) - 3)
        Else
            sCar = Chr(Asc(Mid(sCadena, bPos, 1)) + 3)
        End If
        WPASS = WPASS + sCar
    Next bPos
    DCript2 = WPASS
    Exit Function

SALTO_ERROR:
    DCript2 = Empty
    MsgBox Err.Description, vbCritical, "ECNLIB01 - DCript2"
End Function

'***************************************************************************************************************************************************************************************************************************************
' ComputerName : Devuelve el nombre de la computadora donde se encuentra alojado el aplicativo actual
'***************************************************************************************************************************************************************************************************************************************
Public Function ComputerName() As String
    Dim KeyName$
    Dim keylen&
    Dim iNull

    keylen& = 2000
    KeyName$ = String$(keylen, 0)

    GetcomputerName KeyName$, keylen&

    iNull = InStr(KeyName, Chr(0))
    ComputerName = Mid(KeyName$, 1, iNull - 1)
End Function

'***************************************************************************************************************************************************************************************************************************************
' UserName : USUARIO DE INICIO DE SESION DEL S.O
'***************************************************************************************************************************************************************************************************************************************
Public Function UserName() As String
    Dim Dummy$
    Dim Success&

    Dummy = Space(100)
    Success = GetUserName(Dummy, 100)
    UserName = Trim$(Dummy)
End Function
'***************************************************************************************************************************************************************************************************************************************
' DevuelveResolucionDePantalla : Resolucion de la pantalla
'***************************************************************************************************************************************************************************************************************************************
Public Function DevuelveResolucionDePantalla() As String
    Dim lBits As Long
    Dim lWidth As Long
    Dim lHeight As Long
    Dim hdc As Long

    lBits = GetDeviceCaps(hdc, BITSPIXEL)
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    DevuelveResolucionDePantalla = LTrim(Str(lWidth)) + "x" + Trim(Str(lHeight))
End Function
'***************************************************************************************************************************************************************************************************************************************
' GeneFile : Función que genera un Nombre de archivo Aleatorio
'***************************************************************************************************************************************************************************************************************************************
Function GeneFile(Optional ByVal sPrefijo As String = "ECNW", _
                  Optional ByVal nValor As Integer) As String
    Randomize
    If nValor = 0 Then nValor = 7
    GeneFile = sPrefijo + Left(Trim(Str(Int(Rnd * 10000000))), nValor)
End Function
'***************************************************************************************************************************************************************************************************************************************
' GrabarValoresDeUnFormulario : Graba los datos del formulario y control guardandolos en archivo INI
'***************************************************************************************************************************************************************************************************************************************
Public Sub GrabarValoresDeUnFormulario(ByVal xForm As Object, _
                                       ByVal sRutaINI As String, _
                                       Optional ByVal sNomEmpresa As String = "")
    On Error GoTo MsgVFor

    Dim sCd As Long
    Dim cadena As String
    Dim n As Long
    Dim N2 As Long
    Dim vRegAux As String
    Dim Control As Control
    Dim FilePRN2 As String
    Dim QPag As Long
    Dim cName As String
    Dim iLB As Integer

    n = FreeFile
    N2 = FreeFile
    QPag = 0: N2 = FreeFile + 1
    FilePRN2 = sRutaINI + GeneFile() + ".PRN"

    Open FilePRN2 For Output As #N2
    Open sRutaINI + GK_FILE_PARAM_INI_WIN_PNE For Input As n
    Do While Not EOF(n)
       Line Input #n, cadena
       sCd = InStr(cadena, sNomEmpresa + "[" + xForm.Name + "]")
       If sCd = 0 Then
          Print #N2, cadena
       End If
    Loop


    cadena = sNomEmpresa + "[" + xForm.Name + "][KPropHEIGHT]=" + CStr(xForm.Height)
    Print #N2, cadena
    cadena = sNomEmpresa + "[" + xForm.Name + "][KPropWIDTH]=" + CStr(xForm.Width)
    Print #N2, cadena
    cadena = sNomEmpresa + "[" + xForm.Name + "][KPropTOP]=" + CStr(xForm.Top)
    Print #N2, cadena
    cadena = sNomEmpresa + "[" + xForm.Name + "][KPropLEFT]=" + CStr(xForm.Left)
    Print #N2, cadena
    cadena = sNomEmpresa + "[" + xForm.Name + "][KPropWindowState]=" + CStr(xForm.WindowState)
    Print #N2, cadena



    For Each Control In xForm.Controls
        Select Case TypeName(Control)
           Case "Label"
              If InStr(Control.Name, "Label") = 0 Then
                 vRegAux = CStr(Control.Caption)
              Else
                 vRegAux = ""
              End If
           Case "TextBox", "MaskEdBox"
              vRegAux = Control.Text
           Case "ComboBox"
              vRegAux = CStr(Control.ListIndex)
           Case "OptionButton", "CheckBox"
              If Control.Value Then
                 vRegAux = ".V."
              Else
                 vRegAux = ".F."
              End If
           Case Else
              vRegAux = ""
        End Select
        cName = Trim(Control.Name)
        If TypeName(Control) = "ListBox" Then
            For iLB = 0 To Control.ListCount - 1
                If Control.Selected(iLB) = True Then
                    vRegAux = CStr(iLB)
                    If Not Trim(vRegAux) = "" Then
                       cadena = sNomEmpresa + "[" + xForm.Name + "][" + Trim(cName) + "]=" + vRegAux
                       Print #N2, cadena
                    End If
                End If
            Next
        Else
            If Control.Index <> "" Then
               cName = Trim(Control.Name) + "(" + Trim(Str(Control.Index)) + ")"
               If Not Trim(vRegAux) = "" Then
                   cadena = sNomEmpresa + "[" + xForm.Name + "][" + Trim(cName) + "]=" + vRegAux
                   Print #N2, cadena
                End If
            End If
        End If
        vRegAux = ""
    Next
    Close N2
    Close n
    Kill sRutaINI + GK_FILE_PARAM_INI_WIN_PNE
    FileCopy FilePRN2, sRutaINI + GK_FILE_PARAM_INI_WIN_PNE
    Kill FilePRN2
    Exit Sub
MsgVFor:
    Resume Next
End Sub
'***************************************************************************************************************************************************************************************************************************************
' AsignarValoresDeUnFormulario : Lee el archivo INI correspondiente al formulario asignando los valores guardados para rediseñar los diferentes controles del formulario incluyéndolo
'***************************************************************************************************************************************************************************************************************************************
Public Sub AsignarValoresDeUnFormulario(ByVal xForm As Object, _
                                        ByVal sRutaINI As String, _
                                        Optional ByVal sNomEmpresa As String = "")
    On Error GoTo msgFile

    Dim sCd As Long, cadena As String, n As Long
    Dim iCd As Long, fCd As String, nCtrl As String, nVal As String
    Dim iPosPropFrm As Integer

    Dim Control As Control
    Dim sNomCTRL_FND As String

    n = FreeFile
    Open sRutaINI + GK_FILE_PARAM_INI_WIN_PNE For Input As n
    Do While Not EOF(n)
       Line Input #n, cadena
       sCd = InStr(cadena, sNomEmpresa + "[" + xForm.Name + "]")
       '***************************************************************************************************
       '==> SE BUSCA EL FORMULARIO EN EL INI
       '***************************************************************************************************
       If sCd <> 0 Then
          iPosPropFrm = InStr(cadena, sNomEmpresa + "[" + xForm.Name + "][KProp")
          If iPosPropFrm <> 0 Then
                '******************************************************************************************
                '==> PROPIEDADES DEL FORMULARIO
                '******************************************************************************************
                '--+-------------------------------+--
                '==> HEIGHT
                '--+-------------------------------+--
                iPosPropFrm = InStr(cadena, sNomEmpresa + "[" + xForm.Name + "][KPropHEIGHT]")
                If iPosPropFrm <> 0 Then
                   nVal = Trim(Mid(cadena, InStr(cadena, "=") + 1))
                   xForm.Height = Val(nVal)
                End If
                '--+-------------------------------+--
                '==> WIDTH
                '--+-------------------------------+--
                iPosPropFrm = InStr(cadena, sNomEmpresa + "[" + xForm.Name + "][KPropWIDTH]")
                If iPosPropFrm <> 0 Then
                   nVal = Trim(Mid(cadena, InStr(cadena, "=") + 1))
                   xForm.Width = Val(nVal)
                End If
                '--+-------------------------------+--
                '==> TOP
                '--+-------------------------------+--
                iPosPropFrm = InStr(cadena, sNomEmpresa + "[" + xForm.Name + "][KPropTOP]")
                If iPosPropFrm <> 0 Then
                   nVal = Trim(Mid(cadena, InStr(cadena, "=") + 1))
                   xForm.Top = Val(nVal)
                End If
                '--+-------------------------------+--
                '==> LEFT
                '--+-------------------------------+--
                iPosPropFrm = InStr(cadena, sNomEmpresa + "[" + xForm.Name + "][KPropLEFT]")
                If iPosPropFrm <> 0 Then
                   nVal = Trim(Mid(cadena, InStr(cadena, "=") + 1))
                   xForm.Left = Val(nVal)
                End If
                '--+-------------------------------+--
                '==> WINDOWSSTATE
                '--+-------------------------------+--
                iPosPropFrm = InStr(cadena, sNomEmpresa + "[" + xForm.Name + "][KPropWindowState]")
                If iPosPropFrm <> 0 Then
                   nVal = Trim(Mid(cadena, InStr(cadena, "=") + 1))
                   xForm.WindowState = Val(nVal)
                End If

          Else
                '******************************************************************************************
                '==> SE LEE LOS CONTROLES DEL FORMULARIO
                '******************************************************************************************
                iCd = sCd + Len(sNomEmpresa + "[" + xForm.Name + "]") + 1
                fCd = InStr(cadena, "=") - 1
                nCtrl = Mid(cadena, iCd, (fCd - iCd))
                nVal = Trim(Mid(cadena, InStr(cadena, "=") + 1))
                For Each Control In xForm.Controls 'Encontrar Control
                    sNomCTRL_FND = Trim(Control.Name)
                    If Control.Index <> "" Then
                        sNomCTRL_FND = Trim(Control.Name) + "(" + Trim(Str(Control.Index)) + ")"
                    End If
                    If sNomCTRL_FND = nCtrl Then
                       Select Case TypeName(Control)
                           Case "Label"
                                Control.Caption = nVal
                           Case "TextBox", "MaskEdBox"
                                Control.Text = nVal
                           Case "ComboBox"
                                If Val(nVal) <> -1 Then
                                   If Control.Count > nVal Then
                                      Control.ListIndex = Val(nVal)
                                   End If
                                End If
                           Case "ListBox"
                                If Val(nVal) <> -1 Then
                                   If Control.Count > nVal Then
                                      If Control.Style = 1 Then
                                         Control.Selected(Val(nVal)) = True
                                      Else
                                         Control.ListIndex = Val(nVal)
                                      End If
                                   End If
                                End If
                           Case "OptionButton"
                                If nVal = ".V." Then
                                  Control.Value = True
                                Else
                                  Control.Value = False
                                End If
                           Case "CheckBox"
                                If nVal = ".V." Then
                                   Control.Value = 1
                                Else
                                   Control.Value = 0
                                End If
                       End Select
                    End If
                Next
          End If
       End If
    Loop
    Close n
    Exit Sub
msgFile:
    If Err.Number = 53 Then
       Open sRutaINI + GK_FILE_PARAM_INI_WIN_PNE For Output As #n
       Print #n, "ERU -L-'"
       Close #n
    Else
       Resume Next
    End If
End Sub


'************************************************************************************************************************************************************************************************************************************************************************************************************
' sAlineaTxTdeBoton : PROCEDIMIENTO QUE ALINEA EL TEXTO DE UN BOTON, DONDE xCommandButton ES UN CONTROL COMMANDBUTTON
'************************************************************************************************************************************************************************************************************************************************************************************************************
Public Sub AlinearTxTdeBoton(ByVal opcEruAlineacion As eruAlineacion, ByVal xCommandButton As Object)
    Dim tmpValue&
    Dim Align&
    Dim ret&
    Dim fAlignment&
    Dim BS_RIGHT&

    Select Case opcEruAlineacion
        Case eruCENTRO
            fAlignment& = A_CENTER

        Case eruARRIBA
            fAlignment& = A_TOP
        Case eruARRIBA_IZQUIERDA
            fAlignment& = A_TOPLEFT
        Case eruARRIBA_DERECHA
            fAlignment& = A_TOPRIGHT

        Case eruABAJO
            fAlignment& = A_BOTTOM
        Case eruABAJO_IZQUIERDA
            fAlignment& = A_BOTTOMLEFT
        Case eruABAJO_DERECHA
            fAlignment& = A_BOTTOMRIGHT

        Case eruIZQUIERDA
            fAlignment& = A_LEFT
        Case eruDERECHA
            fAlignment& = A_RIGHT
    End Select

    tmpValue& = GetWindowLong&(xCommandButton.hWnd, GWL_STYLE) And Not BS_RIGHT
    ret& = SetWindowLong&(xCommandButton.hWnd, GWL_STYLE, tmpValue& Or fAlignment&)
    xCommandButton.Refresh
End Sub
'************************************************************************************************************************************************************************************************************************************************************************************************************
' CambiarComboHeight : CAMBIAR EL TAMAÑO(HEIGHT) DE UN COMBO DONDE xComboBox ES UN CONTROL COMBOBOX
'************************************************************************************************************************************************************************************************************************************************************************************************************
Public Sub CambiarComboHeight(ByVal xComboBox As Object, ByVal lNuevoHeight As Long)
    With xComboBox
        Call SendMessage(.hWnd, CB_SETITEMHEIGHT, -1, lNuevoHeight)
    End With
End Sub
'************************************************************************************************************************************************************************************************************************************************************************************************************
' CambiaComboListaAncho : AMPLIAR EL ANCHO DE LA LISTA DE UN COMBO DONDE xComboBox ES UN CONTROL COMBOBOX
'************************************************************************************************************************************************************************************************************************************************************************************************************
Public Sub CambiarComboListaAncho(ByVal xComboBox As Object, ByVal lAncho As Long)
    On Error Resume Next
    SendMessage xComboBox.hWnd, CB_SETDROPPEDWIDTH, lAncho, 0
End Sub
'************************************************************************************************************************************************************************************************************************************************************************************************************
' CambiarComboListaLargo : AMPLIAR EL LARGO DE LA LISTA DE UN COMBO DONDE xComboBox ES UN CONTROL COMBOBOX
'************************************************************************************************************************************************************************************************************************************************************************************************************
Public Sub CambiarComboListaLargo(ByVal xComboBox As Object, ByVal lNuevoLargo As Long)
    Dim iAntiguaScala As Integer

    On Error Resume Next

    If TypeOf xComboBox.Parent Is Frame Then Exit Sub

    iAntiguaScala = xComboBox.Parent.ScaleMode
    xComboBox.Parent.ScaleMode = vbPixels

    MoveWindow xComboBox.hWnd, xComboBox.Left, xComboBox.Top, xComboBox.Width, lNuevoLargo, 1
    xComboBox.Parent.ScaleMode = iAntiguaScala
End Sub
'***************************************************************************************************************************************************************************************************************************************
' DesplegarCb : PROCEDIMIENTO QUE DEPLIEGA O NO UN COMBOBOX DONDE xCbo ES UN CONTROL COMBOBOX
'***************************************************************************************************************************************************************************************************************************************
Public Sub DesplegarCb(ByVal xCb As Object, ByVal bsDespliegaCb As Boolean)
    On Error Resume Next
    SendMessage xCb.hWnd, CB_sDespliegaCb, bsDespliegaCb, 0
End Sub
'***************************************************************************************************************************************************************************************************************************************
' CentrarFrm : PROCEDIMIENTO QUE CENTRA UN FORMULARIO
'***************************************************************************************************************************************************************************************************************************************
Public Sub CentrarFrm(ByVal xForm As Object)
    xForm.Move (Screen.Width - xForm.Width) / 2, (Screen.Height - xForm.Height) / 2
End Sub
'***************************************************************************************************************************************************************************************************************************************
' DeGradarColoRPictuRe : PROCEDIMIENTO QUE REALIZA UN DEGRADE EN UN PICTUREBOX DONDE xPictureBox ES UN CONTROL PICTUREBOX
'***************************************************************************************************************************************************************************************************************************************
Public Sub DeGradarColoRPictuRe(ByVal xPictureBox As Object, _
                                ByVal INICIO_R As Single, _
                                ByVal INICIO_G As Single, _
                                ByVal INICIO_B As Single, _
                                ByVal FIN_R As Single, _
                                ByVal FIN_G As Single, _
                                ByVal FIN_B As Single, _
                                ByVal INICIO As Single, _
                                ByVal FIN As Single, ByVal eruDiReccion As EruDegradaDireccion)
    On Error Resume Next
    Dim HGT As Single
    Dim Punto_X As Single
    Dim Punto_Y As Single
    Dim R As Single
    Dim G As Single
    Dim b As Single
    Dim dR As Single
    Dim dG As Single
    Dim dB As Single
    Dim PUNTO As Single

    xPictureBox.AutoRedraw = True
    Punto_X = xPictureBox.ScaleWidth
    Punto_Y = xPictureBox.ScaleHeight
    HGT = FIN - INICIO
    dR = (FIN_R - INICIO_R) / HGT
    dG = (FIN_G - INICIO_G) / HGT
    dB = (FIN_B - INICIO_B) / HGT
    R = INICIO_R
    G = INICIO_G
    b = INICIO_B
    For PUNTO = INICIO To FIN
        If eruDiReccion = eRuDeGrada_VERTICAL Then xPictureBox.Line (0, PUNTO)-(Punto_X, PUNTO), RGB(R, G, b)
        If eruDiReccion = eRuDeGrada_HORIZONTAL Then xPictureBox.Line (PUNTO, 0)-(PUNTO, Punto_Y), RGB(R, G, b)
        R = R + dR
        G = G + dG
        b = b + dB
    Next PUNTO
    
    xPictureBox.Refresh
End Sub
'***************************************************************************************************************************************************************************************************************************************
' Esperar : PROCEDIMIENTO QUE HACE UNA PAUSA EN EL PROGRAMA POR UN NUMERO DE SEGUNDOS DETERMINADO
'***************************************************************************************************************************************************************************************************************************************
Public Sub Esperar(Optional ByVal iSegundosEspera As Integer = 1, _
                   Optional ByVal blSW_DoEvents As Boolean = False)
    Dim sHoraActual As String
    Dim sHoraDeCese As String
    Dim sHoraSecond As String
    
    sHoraDeCese = FormatDateTime(DateAdd("s", iSegundosEspera, Time), vbLongTime)
    sHoraSecond = FormatDateTime(Time, vbLongTime)
    Do While sHoraDeCese <> sHoraActual
        sHoraActual = FormatDateTime(Time, vbLongTime)
        If sHoraActual <> sHoraSecond Then
            sHoraSecond = sHoraActual
            If blSW_DoEvents = True Then DoEvents
        End If
    Loop
End Sub
'***************************************************************************************************************************************************************************************************************************************
' Esperar : PROCEDIMIENTO QUE HACE UNA PAUSA EN EL PROGRAMA POR UN NUMERO DE SEGUNDOS DETERMINADO
'***************************************************************************************************************************************************************************************************************************************
Public Sub DeshabilitarBotonXdeForm(ByRef xForm As Object)
    Dim hMenu As Long
    Dim nCount As Long

    hMenu = GetSystemMenu(xForm.hWnd, 0)
    nCount = GetMenuItemCount(hMenu)
    Call RemoveMenu(hMenu, _
                    nCount - 1, _
                    MF_DISABLED Or MF_BYPOSITION)
    DrawMenuBar xForm.hWnd
End Sub
'***************************************************************************************************************************************************************************************************************************************
' BuscarEquiposDeRed : UTILIZA EL COMMON DIALOG DEL WINDOWS PARA BUSCAR EQUIPOS DE RED
'***************************************************************************************************************************************************************************************************************************************
Public Function BuscarEquiposDeRed(ByRef xForm As Object) As String
    Dim BI As BROWSEINFO
    Dim pidl As Long
    Dim sPath As String
    Dim Pos As Integer

    ' Obtener el pidl de la carpeta Entorno de red
    If SHGetSpecialFolderLocation(xForm.hWnd, CSIDL_NETWORK, pidl) = ERROR_SUCCESS Then

        With BI
            .hOwner = xForm.hWnd
            .pidlRoot = pidl
            .pszDisplayName = Space$(MAX_PATH)
            .lpszTitle = " Seleccionar el ordenador de la red de la lista :"
            .ulFlags = BIF_BROWSEFORCOMPUTER
        End With

        ' Esto abre el diálogo para buscar el equipo de red
        If SHBrowseForFolder(BI) <> 0 Then
            'retorna el valor a la función, es decir el nombre del equipo
            BuscarEquiposDeRed = Replace(BI.pszDisplayName, Chr(0), vbNullString)
        End If

    End If
End Function
