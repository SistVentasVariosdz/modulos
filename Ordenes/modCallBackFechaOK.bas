Attribute VB_Name = "modCallBackFechaOK"
Option Explicit
Public oCallBack As frmCallBackFechaOK

Private Declare Function EnumSystemLocales Lib "kernel32" Alias "EnumSystemLocalesA" (ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long
Private Declare Function EnumCalendarInfo Lib "kernel32" Alias "EnumCalendarInfoA" (ByVal lpCalInfoEnumProc As Long, ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long) As Long
Private Declare Function EnumDateFormats Lib "kernel32" Alias "EnumDateFormatsA" (ByVal lpDateFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetCurrencyFormatBynum Lib "kernel32" Alias "GetCurrencyFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As String, ByVal lpFormat As Long, ByVal lpCurrencyStr As String, ByVal cchCurrency As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const LCID_SUPPORTED = 2
Private Const LOCALE_SYSTEM_DEFAULT = &H800
Private Const ENUM_ALL_CALENDARS = &HFFFF                '  enumerate all calendars
Private Const DATE_LONGDATE = &H1 ' &H2         '  use long date picture

Private Const CAL_ICALINTVALUE = &H1                     '  calendar type
Private Const CAL_SCALNAME = &H2                         '  native name of calendar
Private Const CAL_IYEAROFFSETRANGE = &H3                 '  starting years of eras
Private Const CAL_SERASTRING = &H4                       '  era name for IYearOffsetRanges
Private Const CAL_SSHORTDATE = &H5                       '  Integer date format string
Private Const CAL_SLONGDATE = &H6                        '  long date format string
Private Const CAL_SDAYNAME1 = &H7                        '  native name for Monday
Private Const CAL_SDAYNAME2 = &H8                        '  native name for Tuesday
Private Const CAL_SDAYNAME3 = &H9                        '  native name for Wednesday
Private Const CAL_SDAYNAME4 = &HA                        '  native name for Thursday
Private Const CAL_SDAYNAME5 = &HB                        '  native name for Friday
Private Const CAL_SDAYNAME6 = &HC                        '  native name for Saturday
Private Const CAL_SDAYNAME7 = &HD                        '  native name for Sunday
Private Const CAL_SABBREVDAYNAME1 = &HE                  '  abbreviated name for Monday
Private Const CAL_SABBREVDAYNAME2 = &HF                  '  abbreviated name for Tuesday
Private Const CAL_SABBREVDAYNAME3 = &H10                 '  abbreviated name for Wednesday
Private Const CAL_SABBREVDAYNAME4 = &H11                 '  abbreviated name for Thursday
Private Const CAL_SABBREVDAYNAME5 = &H12                 '  abbreviated name for Friday
Private Const CAL_SABBREVDAYNAME6 = &H13                 '  abbreviated name for Saturday
Private Const CAL_SABBREVDAYNAME7 = &H14                 '  abbreviated name for Sunday
Private Const CAL_SMONTHNAME1 = &H15                     '  native name for January
Private Const CAL_SMONTHNAME2 = &H16                     '  native name for February
Private Const CAL_SMONTHNAME3 = &H17                     '  native name for March
Private Const CAL_SMONTHNAME4 = &H18                     '  native name for April
Private Const CAL_SMONTHNAME5 = &H19                     '  native name for May
Private Const CAL_SMONTHNAME6 = &H1A                     '  native name for June
Private Const CAL_SMONTHNAME7 = &H1B                     '  native name for July
Private Const CAL_SMONTHNAME8 = &H1C                     '  native name for August
Private Const CAL_SMONTHNAME9 = &H1D                     '  native name for September
Private Const CAL_SMONTHNAME10 = &H1E                    '  native name for October
Private Const CAL_SMONTHNAME11 = &H1F                    '  native name for November
Private Const CAL_SMONTHNAME12 = &H20                    '  native name for December
Private Const CAL_SMONTHNAME13 = &H21                    '  native name for 13th month (if any)
Private Const CAL_SABBREVMONTHNAME1 = &H22               '  abbreviated name for January
Private Const CAL_SABBREVMONTHNAME2 = &H23               '  abbreviated name for February
Private Const CAL_SABBREVMONTHNAME3 = &H24               '  abbreviated name for March
Private Const CAL_SABBREVMONTHNAME4 = &H25               '  abbreviated name for April
Private Const CAL_SABBREVMONTHNAME5 = &H26               '  abbreviated name for May
Private Const CAL_SABBREVMONTHNAME6 = &H27               '  abbreviated name for June
Private Const CAL_SABBREVMONTHNAME7 = &H28               '  abbreviated name for July
Private Const CAL_SABBREVMONTHNAME8 = &H29               '  abbreviated name for August
Private Const CAL_SABBREVMONTHNAME9 = &H2A               '  abbreviated name for September
Private Const CAL_SABBREVMONTHNAME10 = &H2B              '  abbreviated name for October
Private Const CAL_SABBREVMONTHNAME11 = &H2C              '  abbreviated name for November
Private Const CAL_SABBREVMONTHNAME12 = &H2D              '  abbreviated name for December
Private Const CAL_SABBREVMONTHNAME13 = &H2E              '  abbreviated name for 13th month (if any)


Public Function Callback1_EnumFormats(ByVal lpstr As Long) As Long
    oCallBack.List1.AddItem agGetStringFromPointer(lpstr)    ' Note string extraction
    Callback1_EnumFormats = 1
End Function


Public Function FechaOK(ByRef dFecha As Date) As String
    Dim dl&
    Dim selloc&
    Dim resbuf$
    Dim lLocale As Long
    Dim i As Integer
    Dim nInSTR As Integer
    Dim sFecha As String
    
    Set oCallBack = New frmCallBackFechaOK
    
    lLocale = GetUserDefaultLCID()
        
    oCallBack.List1.Clear
    selloc& = lLocale
        
    
    dl& = EnumDateFormats(AddressOf Callback1_EnumFormats, selloc, DATE_LONGDATE)
    resbuf$ = String$(50, 0)
    If dl& > 0 Then oCallBack.List1.AddItem Left$(resbuf$, dl&)
        
    For i = 0 To oCallBack.List1.ListCount - 1
        nInSTR = InStr(UCase(oCallBack.List1.List(i)), UCase(kFORMAT_TO_PRINT))
        If nInSTR = 0 Then
            nInSTR = InStr(UCase(oCallBack.List1.List(i)), UCase(kFORMAT_TO_PRINTSHORT))
        End If
        If nInSTR = 0 Then
            sFecha = Format(dFecha, kFORMAT_TO_PRINTSHORT)
            FechaOK = sFecha
            Exit For
        Else
            sFecha = CStr(dFecha)
            FechaOK = sFecha
        End If
    Next
    Unload oCallBack
    Set oCallBack = Nothing
End Function



