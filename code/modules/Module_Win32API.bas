Attribute VB_Name = "Module_Win32API"
' ========================================================================
' == DanielsXLToolbox   (c) 2008-2013 Daniel Kraus   Licensed under GPLv2
' ========================================================================
' == DanielsXLToolbox.Module_API
' ==
' ==
' ==
' ==
' == 24-Jan-11 preliminarily 64bit safe
' == 05-Nov-11 GetText support
' ==



Option Explicit
Option Private Module



' API constants
    Const VK_SHIFT = &H10
    Const VK_CONTROL = &H11
    Const VK_ALT = &H12
    Const VK_ESCAPE = &H1B
    
    Public Const INVALID_HANDLE_VALUE = -1
    
    Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
    Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY  As Long = &H2000
    Private Const FORMAT_MESSAGE_FROM_HMODULE  As Long = &H800
    Private Const FORMAT_MESSAGE_FROM_STRING  As Long = &H400
    Private Const FORMAT_MESSAGE_FROM_SYSTEM  As Long = &H1000
    Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK  As Long = &HFF
    Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
    Private Const FORMAT_MESSAGE_TEXT_LEN  As Long = &HA0 ' from VC++ ERRORS.H file
    Public Const MAX_PATH = 260
        
    Type rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    
    Type PAINTSTRUCT
            #If VBA7 Then
                hDC As LongPtr
            #Else
                hDC As Long
            #End If
            fErase As Long
            rcPaint As rect
            fRestore As Long
            fIncUpdate As Long
            rgbReserved(32) As Byte 'this was private declared incorrectly in VB API viewer
    End Type
    
    Public Type tCursorPos
        x As Long
        y As Long
    End Type

    Type DWORDLONG
        Lo As Long
        Hi As Long
    End Type
    
    Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        #If VBA7 Then
            lpMinimumApplicationAddress As LongPtr
            lpMaximumApplicationAddress As LongPtr
            dwActiveProcessorMask As LongPtr
        #Else
            lpMinimumApplicationAddress As Long
            lpMaximumApplicationAddress As Long
            dwActiveProcessorMask As Long
        #End If
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
    End Type
    

    Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF

    Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
    End Type

    Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(0 To 31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
    End Type
    
    ' Predefined Clipboard Formats
    Public Enum CF_ClipboardFormat
        CF_TEXT = 1
        CF_BITMAP = 2
        CF_METAFILEPICT = 3
        CF_SYLK = 4
        CF_DIF = 5
        CF_TIFF = 6
        CF_OEMTEXT = 7
        CF_DIB = 8
        CF_PALETTE = 9
        CF_PENDATA = 10
        CF_RIFF = 11
        CF_WAVE = 12
        CF_UNICODETEXT = 13
        CF_ENHMETAFILE = 14
        CF_OWNERDISPLAY = &H80
        CF_DSPTEXT = &H81
        CF_DSPBITMAP = &H82
        CF_DSPMETAFILEPICT = &H83
        CF_DSPENHMETAFILE = &H8E
    End Enum
    
    
    ' ================================================================================
    ' = From Microsoft: http://support.microsoft.com/default.aspx?scid=kb;en-us;145679
    ' ================================================================================
    
    Public Const REG_SZ As Long = 1
    Public Const REG_DWORD As Long = 4
    
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    
    Public Const ERROR_NONE = 0
    Public Const ERROR_BADDB = 1
    Public Const ERROR_BADKEY = 2
    Public Const ERROR_CANTOPEN = 3
    Public Const ERROR_CANTREAD = 4
    Public Const ERROR_CANTWRITE = 5
    Public Const ERROR_OUTOFMEMORY = 6
    Public Const ERROR_ARENA_TRASHED = 7
    Public Const ERROR_ACCESS_DENIED = 8
    Public Const ERROR_INVALID_PARAMETERS = 87
    Public Const ERROR_NO_MORE_ITEMS = 259
    
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    Public Const KEY_ALL_ACCESS = &H3F
    
    Public Const REG_OPTION_NON_VOLATILE = 0
    

    Private Const SE_PRIVILEGE_ENABLED As Long = &H2
    Private Const SE_RESTORE_NAME As String = "SeRestorePrivilege"
    Private Const SE_BACKUP_NAME As String = "SeBackupPrivilege"
    
    Private Const TOKEN_QUERY As Long = &H8&
    Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
    
    Private Const REG_FORCE_RESTORE As Long = 8& 'allows restore to open key
    
    Private Type LUID
        lowpart As Long
        highpart As Long
    End Type
    
    Private Type LUID_AND_ATTRIBUTES
        pLuid As LUID
        Attributes As Long
    End Type
    
    Private Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        Privileges As LUID_AND_ATTRIBUTES
    End Type
    
        
    '   OS                      Version
    '   ======================  =======
    '   Windows 7               6.1
    '   Windows Server 2008 R2  6.1
    '   Windows Server 2008     6.0
    '   Windows Vista           6.0
    '   Windows Server 2003 R2  5.2
    '   Windows Server 2003     5.2
    '   Windows XP              5.1
    '   Windows 2000            5.0
    
    Type tOSVERSIONINFO
        OSVersionInfoSize As Long
        MajorVersion As Long
        MinorVersion As Long
        BuildNumber As Long
        PlatformID As Long
        
        VersionString(1 To 128) As Byte
    End Type
    
    
#If VBA7 Then
    Private Declare PtrSafe Function SetDllDirectory Lib "kernel32.dll" Alias "SetDllDirectoryA" ( _
        ByVal lpPathName As String) As Boolean
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, _
        ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr) As Long
    Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

        ' GetTickCount:
        ' =============
        '
        ' Ticks are counted in milliseconds from system startup;
        ' because it is a Long value, it will overflow:
        ' An unsigned Long integer will overflow after 49.7 days;
        ' a signed Long integer (max. value &h7FFFFFFF) will
        ' overflow/change sign after
        ' &h7FFFFFFF / 24 [h/d] /60 [min/h] /60 [s/min]/1000 [msec/s]
        ' = 24.855 days.
    
    Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Declare PtrSafe Sub EmptyClipboard Lib "user32" ()
    Declare PtrSafe Function OpenClipboard Lib "user32" ( _
        ByVal hWndNewOwner As LongPtr) As Long
    Declare PtrSafe Function GetClipboardData Lib "user32" ( _
        ByVal uFormat As CF_ClipboardFormat) As LongPtr
    Declare PtrSafe Function SetClipboardData Lib "user32" ( _
        ByVal uFormat As CF_ClipboardFormat, ByRef hMEM As String) As LongPtr
    Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr) As Long
    Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
        ByVal HKEY As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As LongPtr) As Long
    Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As LongPtr) As Long
    Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
        ByVal HKEY As LongPtr, ByVal lpSubKey As String, ByVal Reserved As Long, _
        ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
        ByVal lpSecurityAttributes As LongPtr, phkResult As LongPtr, lpdwDisposition As Long) As Long
    Declare PtrSafe Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal HKEY As LongPtr, ByVal lpValueName As String, _
        ByVal lpReserved As LongPtr, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    Declare PtrSafe Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal HKEY As LongPtr, ByVal lpValueName As String, _
        ByVal lpReserved As LongPtr, lpType As Long, _
        lpData As LongPtr, lpcbData As Long) As Long
    Declare PtrSafe Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
        ByVal HKEY As LongPtr, ByVal lpValueName As String, _
        ByVal lpReserved As LongPtr, lpType As Long, _
        ByVal lpData As LongPtr, lpcbData As Long) As Long
    Declare PtrSafe Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" ( _
        ByVal HKEY As LongPtr, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
        ByVal lpValue As String, ByVal cbData As Long) As Long
    Declare PtrSafe Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" ( _
        ByVal HKEY As LongPtr, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
        lpValue As LongPtr, ByVal cbData As Long) As Long
    Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
        ByVal dwFlags As Long, ByVal lpSource As Any, ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
        ByRef Arguments As LongPtr) As Long
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, _
        ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
        lpPoint As tCursorPos) As Long
    Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" ( _
        lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    Private Declare PtrSafe Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" ( _
        ByVal HKEY As LongPtr, ByVal lpFile As String, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" ( _
        ByVal HKEY As LongPtr, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
    Private Declare PtrSafe Function RegSaveKeyEx Lib "advapi32.dll" Alias "RegSaveKeyExA" ( _
        ByVal HKEY As LongPtr, ByVal lpFile As String, lpSecurityAttributes As Any, ByVal flags As Integer) As Long
    Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
    Private Declare PtrSafe Function OpenProcessToken Lib "advapi32.dll" ( _
        ByVal ProcessHandle As LongPtr, ByVal DesiredAccess As Long, TokenHandle As LongPtr) As Long
    Private Declare PtrSafe Function AdjustTokenPrivileges Lib "advapi32.dll" ( _
        ByVal TokenHandle As LongPtr, ByVal DisableAllPrivileges As Long, _
        NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
        PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
    Private Declare PtrSafe Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" ( _
        ByVal lpSystemName As Any, ByVal lpName As String, ByRef lpLuid As LUID) As Long
    Private Declare PtrSafe Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" _
        (VersionInfo As tOSVERSIONINFO) As Long
    Private Declare PtrSafe Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" ( _
        ByVal lpString As LongPtr) As Long
#Else
    Private Declare Function SetDllDirectory Lib "kernel32.dll" Alias "SetDllDirectoryA" (ByVal lpPathName As String) As Boolean
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

        ' GetTickCount:
        ' =============
        '
        ' Ticks are counted in milliseconds from system startup;
        ' because it is a Long value, it will overflow:
        ' An unsigned Long integer will overflow after 49.7 days;
        ' a signed Long integer (max. value &h7FFFFFFF) will
        ' overflow/change sign after
        ' &h7FFFFFFF / 24 [h/d] /60 [min/h] /60 [s/min]/1000 [msec/s]
        ' = 24.855 days.
    
    Declare Function GetTickCount Lib "kernel32" () As Long
    Declare Sub EmptyClipboard Lib "user32" ()
    Declare Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As Long) As Long
    Declare Function GetClipboardData Lib "user32" (ByVal uFormat As CF_ClipboardFormat) As Long
    Declare Function SetClipboardData Lib "user32" (ByVal uFormat As CF_ClipboardFormat, ByRef hMEM As String) As Long
    Declare Function CloseClipboard Lib "user32" () As Long
    Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKEY As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long
    Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal HKEY As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
    Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
    Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
    Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
    Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As tCursorPos) As Long
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal HKEY As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
    Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal HKEY As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
    Private Declare Function RegSaveKeyEx Lib "advapi32.dll" Alias "RegSaveKeyExA" (ByVal HKEY As Long, _
        ByVal lpFile As String, lpSecurityAttributes As Any, ByVal flags As Integer) As Long
    Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
    Private Declare Function OpenProcessToken Lib "advapi32.dll" _
        (ByVal ProcessHandle As Long, _
        ByVal DesiredAccess As Long, _
        TokenHandle As Long) As Long
    Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
        (ByVal TokenHandle As Long, _
        ByVal DisableAllPrivileges As Long, _
        NewState As TOKEN_PRIVILEGES, _
        ByVal BufferLength As Long, _
        PreviousState As TOKEN_PRIVILEGES, _
        ReturnLength As Long) As Long
    Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" _
        (ByVal lpSystemName As Any, _
        ByVal lpName As String, _
        lpLuid As LUID) As Long
    Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" _
        (VersionInfo As tOSVERSIONINFO) As Long
    Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" ( _
        ByVal lpString As Long) As Long
#End If
    




Public Function GetSystemErrorMessageText(ErrorNumber As Long) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GetSystemErrorMessageText
    ' -------------------------
    ' By Chp Pearson, www.cpearson.com, chip@cpearson.com
    ' See www.cpearson.com/Excel/FormatMessage.aspx for
    ' additional information.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GetSystemErrorMessageText
    '
    ' This function gets the system error message text that corresponds
    ' to the error code parameter ErrorNumber. This value is the value returned
    ' by Err.LastDLLError or by GetLastError, or occasionally as the returned
    ' result of a Windows API function.
    '
    ' These are NOT the error numbers returned by Err.Number (for these
    ' errors, use Err.Description to get the description of the error).
    '
    ' In general, you should use Err.LastDllError rather than GetLastError
    ' because under some circumstances the value of GetLastError will be
    ' reset to 0 before the value is returned to VBA. Err.LastDllError will
    ' always reliably return the last error number raised in an API function.
    '
    ' The function returns vbNullString is an error occurred or if there is
    ' no error text for the specified error number.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim ErrorText As String
    Dim TextLen As Long
    Dim FormatMessageResult As Long
    Dim LangID As Long
    
    ''''''''''''''''''''''''''''''''
    ' Initialize the variables
    ''''''''''''''''''''''''''''''''
    LangID = 0&   ' Default language
    ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, vbNullChar)
    TextLen = FORMAT_MESSAGE_TEXT_LEN
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call FormatMessage to get the text of the error message text
    ' associated with ErrorNumber.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    FormatMessageResult = FormatMessage( _
                            dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or _
                                     FORMAT_MESSAGE_IGNORE_INSERTS, _
                            lpSource:=0&, _
                            dwMessageId:=ErrorNumber, _
                            dwLanguageId:=LangID, _
                            lpBuffer:=ErrorText, _
                            nSize:=TextLen, _
                            Arguments:=0&)
    
    If FormatMessageResult = 0& Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' An error occured. Display the error number, but
        ' don't call GetSystemErrorMessageText to get the
        ' text, which would likely cause the error again,
        ' getting us into a loop.
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        MsgBox t("An error occurred with the FormatMessage API function call.\n" & _
               "Error: {} (Hex {}).", CStr(Err.LastDllError), Hex(Err.LastDllError))
        GetSystemErrorMessageText = t("An internal system error occurred with the\n" & _
            "FormatMessage API function: {}.\n" & _
            "No futher information is available.", CStr(Err.LastDllError))
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If FormatMessageResult is not zero, it is the number
    ' of characters placed in the ErrorText variable.
    ' Take the left FormatMessageResult characters and
    ' return that text.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ErrorText = Left$(ErrorText, FormatMessageResult)
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Get rid of the trailing vbCrLf, if present.
    '''''''''''''''''''''''''''''''''''''''''''''
    If Len(ErrorText) >= 2 Then
        If Right$(ErrorText, 2) = vbCrLf Then
            ErrorText = Left$(ErrorText, Len(ErrorText) - 2)
        End If
    End If
    
    ''''''''''''''''''''''''''''''''
    ' Return the error text as the
    ' result.
    ''''''''''''''''''''''''''''''''
    GetSystemErrorMessageText = ErrorText

End Function


Function ShiftKeyPressed() As Boolean
    ShiftKeyPressed = (GetKeyState(VK_SHIFT) < 0)
End Function


Function ShiftAltOrCtrlKeyPressed() As Boolean
    ShiftAltOrCtrlKeyPressed = (GetKeyState(VK_SHIFT) < 0) Or _
        (GetKeyState(VK_ALT) < 0) Or _
        (GetKeyState(VK_CONTROL) < 0)
    
End Function

Sub DownloadFile(aURL As String, aPath As String)
    On Error Resume Next
    URLDownloadToFile 0, aURL, aPath, 0, 0
End Sub





' ================================================================================
' = From Microsoft: http://support.microsoft.com/default.aspx?scid=kb;en-us;145679
' ================================================================================

#If VBA7 Then
    Public Function RegSetValueEx(ByVal HKEY As LongPtr, sValueName As String, _
       lType As Long, vValue As Variant) As Long
    Dim lValue As LongPtr
#Else
    Public Function RegSetValueEx(ByVal HKEY As Long, sValueName As String, _
       lType As Long, vValue As Variant) As Long
    Dim lValue As Long
#End If
       Dim sValue As String
       Select Case lType
           Case REG_SZ
               sValue = vValue & Chr$(0)
               RegSetValueEx = RegSetValueExString(HKEY, sValueName, 0&, _
                                                lType, sValue, Len(sValue))
           Case REG_DWORD
               lValue = vValue
               RegSetValueEx = RegSetValueExLong(HKEY, sValueName, 0&, _
                                                lType, lValue, 4)
           End Select
End Function


#If VBA7 Then
    Function RegQueryValueEx(ByVal lhKey As LongPtr, ByVal szValueName As _
            String, vValue As Variant) As Long
    Dim lValue As LongPtr
#Else
    Function RegQueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
            String, vValue As Variant) As Long
    Dim lValue As Long
#End If
    Dim lrc As Long
    Dim cch As Long
    Dim lType As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)

             lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
             If lrc = ERROR_NONE Then
                 vValue = Left$(sValue, cch - 1)
             Else
                 vValue = Empty
             End If
        ' For DWORDS
         Case REG_DWORD:
             lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
             If lrc = ERROR_NONE Then vValue = lValue
         Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
       RegQueryValueEx = lrc
       Exit Function

QueryValueExError:
       Resume QueryValueExExit
End Function


#If VBA7 Then
    Sub RegCreateNewKey(sNewKeyName As String, lPredefinedKey As LongPtr)
    Dim hNewKey As LongPtr         'handle to the new key
#Else
    Sub RegCreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
    Dim hNewKey As Long         'handle to the new key
#End If
    Dim lRetVal As Long         'result of the RegCreateKeyEx function

    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
              vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
              0&, hNewKey, lRetVal)
    RegCloseKey hNewKey
End Sub
   
   
Sub RegSetKeyValue(sKeyName As String, sValueName As String, _
vValueSetting As Variant, lValueType As Long)
#If VBA7 Then
    Dim lRetVal As LongPtr      'result of the SetValueEx function
    Dim HKEY As LongPtr         'handle of open key
#Else
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim HKEY As Long         'handle of open key
#End If

    'open the specified key
    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, _
                              KEY_SET_VALUE, HKEY)
    lRetVal = RegSetValueEx(HKEY, sValueName, lValueType, vValueSetting)
    RegCloseKey (HKEY)
End Sub


Function RegQueryValue(sKeyName As String, sValueName As String) As Variant
    Dim lRetVal As Long         'result of the API functions
#If VBA7 Then
    Dim HKEY As LongPtr         'handle of opened key
#Else
    Dim HKEY As Long         'handle of opened key
#End If
    Dim vValue As Variant      'setting of queried value

    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, _
        KEY_QUERY_VALUE, HKEY)
    lRetVal = RegQueryValueEx(HKEY, sValueName, vValue)
    RegQueryValue = vValue
    RegCloseKey HKEY
End Function


Private Function RegEnablePrivilege(seName As String) As Boolean
' From:
' http://www.keyongtech.com/2293686-backup-specific-registry-settings

#If VBA7 Then
    Dim hToken As LongPtr
#Else
    Dim hToken As Long
#End If
    Dim bufflen As Long
    Dim returnlen As Long
    Dim lu As LUID
    Dim tp As TOKEN_PRIVILEGES
    Dim tp_prev As TOKEN_PRIVILEGES

    If OpenProcessToken(GetCurrentProcess(), _
        TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, _
        hToken) <> 0 Then
    
        If LookupPrivilegeValue(0&, seName, lu) <> 0 Then
        
            ' Set it up to adjust the program's security privilege.
            With tp
            .PrivilegeCount = 1
            .Privileges.Attributes = SE_PRIVILEGE_ENABLED
            .Privileges.pLuid = lu
            End With
            
            bufflen = Len(tp_prev)
            
            RegEnablePrivilege = (AdjustTokenPrivileges(hToken, _
                False, _
                tp, _
                bufflen, _
                tp_prev, _
                returnlen) <> 0)
        
        End If 'LookupPrivilegeValue
    
    End If 'OpenProcessToken

End Function


#If VBA7 Then
    Function SaveRegistryKey(HKEY As LongPtr, sKeyName As String, sFileName As String) As Boolean
    Dim Handle As LongPtr
#Else
    Function SaveRegistryKey(HKEY As Long, sKeyName As String, sFileName As String) As Boolean
    Dim Handle As Long
#End If
' Save a registry key, its values and all subkey and their values
' to a file.
    Dim res As Long

    If RegEnablePrivilege(SE_BACKUP_NAME) Then
        If RegOpenKeyEx(HKEY, sKeyName, 0, KEY_ALL_ACCESS, Handle) = 0 Then
            res = RegSaveKeyEx(Handle, sFileName, ByVal 0, 1) ' 4 = REG_NO_COMPRESSION
            ' ErrorHandler.Log "RegSaveKey(" & Handle & ", " & sFileName & ", 0) = " & res & " (" & GetSystemErrorMessageText(res) & ")"
            SaveRegistryKey = (res = 0)
            RegCloseKey Handle
        ' Else
        '     ErrorHandler.Log "Could not open key: " & sKeyName
        End If
    ' Else
    '     ErrorHandler.Log "Could not obtain privileges to export registry key."
    End If
End Function




    
' ==================================================================
' == Clipboard wrappers
' ==================================================================


Sub OpenWindowsClipboard()
    OpenClipboard ApphWnd
End Sub

Sub TextToClipboard(ByRef s As String)
    Dim d As New DataObject
    With d
        .SetText s
        .PutInClipboard
    End With
End Sub


Function AddDLLPath(ByRef path As String) As Boolean
    Dim s As String
    s = path & Chr(0)
    AddDLLPath = SetDllDirectory(s)
End Function


Function RemoveDLLPath() As Boolean
    RemoveDLLPath = AddDLLPath("")
End Function

Function TimeZoneOffsetStr() As String
' Returns the time zone offset string
' in the form "-0400" or "+0100"

    Dim tz As TIME_ZONE_INFORMATION
    Dim sign As String
    
    If GetTimeZoneInformation(tz) <> TIME_ZONE_ID_INVALID Then
        With tz
            If .Bias < 0 Then sign = "+" Else sign = "-"
            
            ' Changed for version 2.71: use Abs(tz.bias) to prevent "+-" to occur
            TimeZoneOffsetStr = _
                sign & _
                Format(Abs(tz.Bias) \ 60, "00") + Format(Abs(tz.Bias) Mod 60, "00")
        End With
    End If
End Function


#If VBA7 Then
    Function Win32Timer(TimerProc As LongPtr, milliseconds As Long) As LongPtr
#Else
    Function Win32Timer(TimerProc As Long, milliseconds As Long) As Long
#End If
    Win32Timer = SetTimer(0, 0, milliseconds, TimerProc)
End Function

#If VBA7 Then
    Function Win32ResetTimer(Timer As LongPtr, TimerProc As LongPtr, milliseconds As Long) As LongPtr
#Else
    Function Win32ResetTimer(Timer As Long, TimerProc As Long, milliseconds As Long) As Long
#End If
    Win32ResetTimer = SetTimer(0, Timer, milliseconds, TimerProc)
End Function

#If VBA7 Then
    Function Win32KillTimer(Timer As LongPtr) As Boolean
#Else
    Function Win32KillTimer(Timer As Long) As Boolean
#End If
    Win32KillTimer = (KillTimer(0, Timer))
End Function


#If VBA7 Then
    Function GetFuncPtr(ByRef lngFnPtr) As LongPtr
        'wrapper function to allow AddressOf to be used within VB
        GetFuncPtr = lngFnPtr
    End Function
#Else
    Function GetFuncPtr(ByRef lngFnPtr) As Long
        'wrapper function to allow AddressOf to be used within VB
        GetFuncPtr = lngFnPtr
    End Function
#End If


Function GetOperatingSystemInfo() As tOSVERSIONINFO
    Dim info As tOSVERSIONINFO
    
    info.OSVersionInfoSize = Len(info)
    If GetVersionEx(info) Then GetOperatingSystemInfo = info
End Function


#If VBA7 Then
    Function PointerToString(ByRef lPtr As LongPtr) As String
#Else
    Function PointerToString(ByRef lPtr As Long) As String
#End If
    ' Adapted from a FreeImage function

    Dim abBuffer() As Byte
    Dim lLength As Long

    If (lPtr) Then
        ' get the length of the ANSI string pointed to by lPtr
        lLength = lstrlen(lPtr) ' Calls the Win32 API
        If (lLength) Then
           ' copy characters to a byte array
           ReDim abBuffer(lLength - 1)
           Call CopyMemoryRead(abBuffer(0), ByVal lPtr, lLength)
           ' convert from byte array to unicode BSTR
           PointerToString = StrConv(abBuffer, vbUnicode)
        End If
    End If

End Function

