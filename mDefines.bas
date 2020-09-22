Attribute VB_Name = "mDefines"
Option Explicit

' Versions:
' VER_PLATFORM_WIN32_WINDOWS(1)
'   W95 = 4.0
'   W98 = 4.1
'   WME = 4.9
' VER_PLATFORM_WIN32_NT(2)
'   WNT Serv = 4.0
'   W2k Prof = 5.0
'   WXP Prof = 5.1

Public Const OS_UNKNOWN = -1
Public Const OS_WIN95 = 0
Public Const OS_WIN98 = 1
Public Const OS_WINNT35 = 2
Public Const OS_WINNT4 = 3
Public Const OS_WIN2K = 4

'Version structure
Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

'dwPlatformId defines
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function LoadLibraryEx Lib "kernel32" _
    Alias "LoadLibraryExA" _
    (ByVal lpLibFileName As String, ByVal hFile As Long, _
     ByVal dwFlags As Long) As Long

Public Declare Function FreeLibrary Lib "kernel32" _
    (ByVal hLibModule As Long) As Long

Public Declare Function LoadLibrary Lib "kernel32" _
    Alias "LoadLibraryA" _
    (ByVal lpLibFileName As String) As Long

Public Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

'*******************************************************************
'** TOC format
'*******************************************************************
Type TOC_TRACK
    Rsvd1 As Byte
    ADR As Byte
    Track As Byte
    Rsvd2 As Byte
    Addr(3) As Byte
End Type

Type TOC
    TocLen(1) As Byte
    FirstTrack As Byte
    LastTrack As Byte
    TocTrack(99) As TOC_TRACK
End Type


