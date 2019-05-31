Attribute VB_Name = "APIGuide32"
Option Explicit

' ------------------------------------------------------------------------
'
'     APIGID32.BAS -- APIGID32.DLL API Declarations for Visual Basic
'
'                       Copyright (C) 1992-1996 Desaware
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that Desaware and Ziff-Davis Press has no
'  warranty, obligation or liability for its contents.
'  Refer to the Ziff-Davis Visual Basic Programmer's Guide to the
'  Win32 API for further information.
'
' ------------------------------------------------------------------------
Type POINTS

    X  As Integer
    y  As Integer

End Type

Private Type FILETIME

    dwLowDateTime As Long
    dwHighDateTime As Long

End Type

#If Win32 Then
    Declare Function GetUserDefaultLCID& Lib "Kernel32" ()
    Declare Function agGetInstance& Lib "apigid32.dll" ()
    Declare Function agPOINTStoLong& Lib "apigid32.dll" (pt As POINTS)
    Declare Sub agCopyData Lib "apigid32.dll" (source As Any, dest As Any, ByVal nCount&)
    Declare Sub agCopyDataBynum _
            Lib "apigid32.dll" _
            Alias "agCopyData" (ByVal source&, _
                                ByVal dest&, _
                                ByVal nCount&)
    Declare Function agGetAddressForObject& Lib "apigid32.dll" (object As Any)
    Declare Function agGetAddressForInteger& _
            Lib "apigid32.dll" _
            Alias "agGetAddressForObject" (intnum%)
    Declare Function agGetAddressForLong& _
            Lib "apigid32.dll" _
            Alias "agGetAddressForObject" (intnum&)
    Declare Function agGetAddressForLPSTR& _
            Lib "apigid32.dll" _
            Alias "agGetAddressForObject" (ByVal lpstring$) ' See warning!
    Declare Function agGetAddressForVBString& Lib "apigid32.dll" (vbstring$)
    Declare Function agGetStringFrom2NullBuffer$ Lib "apigid32.dll" (ByVal ptr&)
    Declare Function agGetStringFromLPSTR$ Lib "apigid32.dll" (ByVal src$)
    Declare Function agGetStringFromPointer$ _
            Lib "apigid32.dll" _
            Alias "agGetStringFromLPSTR" (ByVal ptr&)
    Declare Function agSwapBytes% Lib "apigid32.dll" (ByVal src%)
    Declare Function agSwapWords& Lib "apigid32.dll" (ByVal src&)
    Declare Function agMakeROP4& Lib "apigid32.dll" (ByVal foreground&, ByVal background&)
    Declare Function agGetWndInstance& Lib "apigid32.dll" (ByVal hwnd&)
    Declare Function agDWORDto2Integers& Lib "apigid32.dll" (ByVal l&, lw%, lh%)
    Declare Function agIsValidName& Lib "apigid32.dll" (ByVal o As Object, ByVal lpname$)
    Declare Function agInp% Lib "apigid32.dll" (ByVal portid%)
    Declare Function agInpw% Lib "apigid32.dll" (ByVal portid%)
    Declare Function agInpd& Lib "apigid32.dll" (ByVal portid%)
    Declare Sub agOutp Lib "apigid32.dll" (ByVal portid%, ByVal outval%)
    Declare Sub agOutpw Lib "apigid32.dll" (ByVal portid%, ByVal outval%)
    Declare Sub agOutpd Lib "apigid32.dll" (ByVal portid%, ByVal outval&)

    ' Declared As Any to allow it to be used within classes, not to mention by other
    ' double long structures
    Declare Sub agSubtractFileTimes Lib "apigid32.dll" (f1 As Any, f2 As Any, f3 As Any)
    Declare Sub agAddFileTimes Lib "apigid32.dll" (f1 As Any, f2 As Any, f3 As Any)
    Declare Sub agNegateFileTime Lib "apigid32.dll" (f1 As Any)
    Declare Function agConvertFileTimeToDouble Lib "apigid32.dll" (f1 As Any) As Double
    Declare Sub agConvertDoubleToFileTime _
            Lib "apigid32.dll" (ByVal d As Double, _
                                f1 As Any)

#Else
    ' Note, not all 16 bit declarations have equivalent 32 bit functions
    ' and vice versa. Nor is their behavior always identical.
    ' Refer to the Visual Basic Programmer's Guide to the Windows API (16 bit)
    ' for documentation on the following functions

    Global Const CTLFLG_USESPALETTE% = 2

    Global Const CTLFLG_HASPALETTE% = 1

    Declare Function agGetControlHwnd% Lib "Apiguide.dll" (hctl As Control)
    Declare Function agGetInstance% Lib "Apiguide.dll" ()
    Declare Sub agCopyData Lib "Apiguide.dll" (source As Any, dest As Any, ByVal nCount%)
    Declare Sub agCopyDataBynum _
            Lib "Apiguide.dll" _
            Alias "agCopyData" (ByVal source&, _
                                ByVal dest&, _
                                ByVal nCount%)
    Declare Function agGetAddressForObject& Lib "Apiguide.dll" (object As Any)
    Declare Function agGetAddressForInteger& _
            Lib "Apiguide.dll" _
            Alias "agGetAddressForObject" (intnum%)
    Declare Function agGetAddressForLong& _
            Lib "Apiguide.dll" _
            Alias "agGetAddressForObject" (intnum&)
    Declare Function agGetAddressForLPSTR& _
            Lib "Apiguide.dll" _
            Alias "agGetAddressForObject" (ByVal lpstring$)
    Declare Function agGetAddressForVBString& Lib "Apiguide.dll" (vbstring$)
    Declare Function agGetStringFromLPSTR$ Lib "Apiguide.dll" (ByVal lpstring$)
    Declare Function agGetControlName$ Lib "Apiguide.dll" (ByVal hwnd%)
    Declare Function agPOINTAPItoLong& Lib "Apiguide.dll" (pt As POINTAPI)
    Declare Function agPOINTStoLong& _
            Lib "Apiguide.dll" _
            Alias "agPOINTAPItoLong" (pt As POINTS)
    Declare Sub agDWORDto2Integers Lib "Apiguide.dll" (ByVal l&, lw%, lh%)
    Declare Function agXPixelsToTwips& Lib "Apiguide.dll" (ByVal pixels%)
    Declare Function agYPixelsToTwips& Lib "Apiguide.dll" (ByVal pixels%)
    Declare Function agXTwipsToPixels% Lib "Apiguide.dll" (ByVal twips&)
    Declare Function agYTwipsToPixels% Lib "Apiguide.dll" (ByVal twips&)
    Declare Function agDeviceCapabilities& _
            Lib "Apiguide.dll" (ByVal hlib%, _
                                ByVal lpszDevice$, _
                                ByVal lpszPort$, _
                                ByVal fwCapability%, _
                                ByVal lpszOutput&, _
                                ByVal lpdm&)
    Declare Function agDeviceMode% _
            Lib "Apiguide.dll" (ByVal hwnd%, _
                                ByVal hModule%, _
                                ByVal lpszDevice$, _
                                ByVal lpszOutput$)
    Declare Function agExtDeviceMode% _
            Lib "Apiguide.dll" (ByVal hwnd%, _
                                ByVal hDriver%, _
                                ByVal lpdmOutput&, _
                                ByVal lpszDevice$, _
                                ByVal lpszPort$, _
                                ByVal lpdmInput&, _
                                ByVal lpszProfile&, _
                                ByVal fwMode%)
    Declare Function agInp% Lib "Apiguide.dll" (ByVal portid%)
    Declare Function agInpw% Lib "Apiguide.dll" (ByVal portid%)
    Declare Sub agOutp Lib "Apiguide.dll" (ByVal portid%, ByVal outval%)
    Declare Sub agOutpw Lib "Apiguide.dll" (ByVal portid%, ByVal outval%)
    Declare Function agHugeOffset& Lib "Apiguide.dll" (ByVal addr&, ByVal offset&)
    Declare Function agVBGetVersion% Lib "Apiguide.dll" ()
    Declare Function agVBSendControlMsg& _
            Lib "Apiguide.dll" (ctl As Control, _
                                ByVal msg%, _
                                ByVal wp%, _
                                ByVal lp&)
    Declare Function agVBSetControlFlags& _
            Lib "Apiguide.dll" (ctl As Control, _
                                ByVal mask&, _
                                ByVal value&)
    Declare Sub agVBScreenToClient Lib "Apiguide.dll" (ctl As Control, pap As POINTS)
    Declare Sub agVBClientToScreen Lib "Apiguide.dll" (ctl As Control, pap As POINTS)
    Declare Function dwVBSetControlFlags& _
            Lib "Apiguide.dll" (ctl As Control, _
                                ByVal mask&, _
                                ByVal value&)

#End If

