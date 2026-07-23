VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SNATIVELANGNAME = &H4
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal lLocale As Long, ByVal lLocaleType As Long, ByVal sLCData As String, ByVal lBufferLength As Long) As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Public Property Get LanguageName() As String
    'Returns the Native Language Name eg. 'English'
    LanguageName = GetLocaleString(LOCALE_SNATIVELANGNAME)
End Property
Private Function GetLocaleString(ByVal lLocaleNum As Long) As String
    'Generic routine to get the locale string from the Operating system.
    Dim lBuffSize As String
    Dim sBuffer As String
    Dim lRet As Long

    lBuffSize = 256
    sBuffer = String$(lBuffSize, vbNullChar)

    'Get the information from the registry
    lRet = GetLocaleInfo(LOCALE_USER_DEFAULT, lLocaleNum, sBuffer, lBuffSize)
    'If lRet > 0 then success - lret is the size of the string returned
    If lRet > 0 Then
        GetLocaleString = Left$(sBuffer, lRet - 1)
    End If
End Function
Public Property Get SysLanguageID() As Long
    'Returns the System Language ID for the machine
    SysLanguageID = GetSystemDefaultLangID
End Property

Public Property Get SysLanguageName() As String
    'Returns the System Language Name eg : English (United Kingdom)
    Dim lLangID As Long
    Dim sBuffer As String
    Dim lBuffSize As Long
    Dim lRet As Long

    On Error GoTo vbErrorHandler

    lLangID = GetSystemDefaultLangID
    'Setup a buffer to receive the settings
    lBuffSize = 50
    sBuffer = String$(lBuffSize, vbNullChar)
    lRet = VerLanguageName(CLng("&H" & lLangID), sBuffer, lBuffSize)
    If lRet > 0 Then
        SysLanguageName = Left$(sBuffer, lRet)
    End If
Exit Property
vbErrorHandler:
    Err.Raise Err.Number, "CGLocaleInfo GetSysLanguageName", Err.Description
End Property
Public Function GetLanguageName(mid) As String
    'Returns the System Language Name eg : English (United Kingdom)
    Dim lLangID As Long
    Dim sBuffer As String
    Dim lBuffSize As Long
    Dim lRet As Long

    'Setup a buffer to receive the settings
    lBuffSize = 50
    sBuffer = String$(lBuffSize, vbNullChar)
    lRet = VerLanguageName(CLng("&H" & mid), sBuffer, lBuffSize)
    If lRet > 0 Then
        GetLanguageName = Left$(sBuffer, lRet)
    End If
End Function
Public Function IsWin3xx() As Boolean
Dim s As OSVERSIONINFOEX
GetVersionEx s
If s.dwPlatformID = VER_PLATFORM_WIN32s Then
    IsWin3xx = True
Else
    IsWin3xx = False
End If
End Function
Public Function IsWin95OSR2() As Boolean
Dim s As OSVERSIONINFOEX
GetVersionEx s
If s.dwPlatformID = VER_PLATFORM_WIN32s And s.dwMajorVersion = 4 Then
    IsWin95OSR2 = False
    ret = LoadLibrary("kernel32.dll")
    If ret <> 0 Then
        lRet = GetProcAddress(hnd, "GetDiskFreeSpaceExA")
        If lRet <> 0 Then
            Win95OSR2 = True
        End If
        FreeLibrary ret
    End If
Else
    IsWinOSR2 = False
End If
End Function
Public Function IsWin95() As Boolean
Dim s As OSVERSIONINFOEX
GetVersionEx s
If s.dwMajorVersion = 4 And s.dwMinorVersion = 0 And s.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS Then
    IsWin95 = True
Else
    IsWin95 = False
End If
End Function

Public Function IsWin98() As Boolean
Dim s As OSVERSIONINFOEX
GetVersionEx s
If s.dwMajorVersion = 4 And s.dwMinorVersion = 10 And s.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS Then
    IsWin98 = True
Else
    IsWin98 = False
End If
End Function
Public Function IsWinMe() As Boolean
Dim s As OSVERSIONINFOEX
GetVersionEx s
If s.dwMajorVersion = 4 And s.dwMinorVersion = 90 And s.dwPlatformID = VER_PLATFORM_WIN32_WINDOWS Then
    IsWinMe = True
Else
    IsWinMe = False
End If
End Function
Public Function IsWin2000() As Boolean
Dim s As OSVERSIONINFOEX
GetVersionEx s
If s.dwMajorVersion = 5 And s.dwMinorVersion = 0 And s.dwPlatformID = VER_PLATFORM_WIN32_NT Then
    IsWin2000 = True
Else
    IsWin2000 = False
End If
End Function
Public Function IsWinNT() As Boolean
Dim s As OSVERSIONINFOEX
GetVersionEx s
If s.dwPlatformID = VER_PLATFORM_WIN32_NT Then
    IsWinNT = True
Else
    IsWinNT = False
End If
End Function
Public Function IsWinNTless5() As Boolean
Dim s As OSVERSIONINFOEX
GetVersionEx s
If s.dwMajorVersion < 5 And s.dwPlatformID = VER_PLATFORM_WIN32_NT Then
    IsWinNTless5 = True
Else
    IsWinNTless5 = False
End If
End Function
Public Function IsWin5() As Boolean
Dim s As OSVERSIONINFOEX
GetVersionEx s
If s.dwMajorVersion = 5 Then
    IsWin5 = True
Else
    IsWin5 = False
End If
End Function

