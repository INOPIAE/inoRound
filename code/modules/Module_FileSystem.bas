Attribute VB_Name = "Module_FileSystem"
' ========================================================================
' == DanielsXLToolbox   (c) 2008-2013 Daniel Kraus   Licensed under GPLv2
' ========================================================================
' == DanielsXLToolbox.Module_FileSystem
' ==
' == Created: 02-Jul-09 19:14
' ==
' ==
' ==
' == 22-Jan-11 preliminarily 64bit safe
' == 05-Nov-11 GetText support
' ==



Option Explicit
Option Private Module
    Public Const InvalidFileNameChars = "\/:*?""<>|" ' Characters that are not allowed in Windows file names

    ' *** Define Special Folder Constants
    Private Const CSIDL_PROGRAMS = 2                   ' Program Groups Folder
    Private Const CSIDL_PERSONAL = 5                   ' Personal Documents Folder
    Private Const CSIDL_FAVORITES = 6                  ' Favorites Folder
    Private Const CSIDL_STARTUP = 7                    ' Startup Group Folder
    Private Const CSIDL_RECENT = 8                     ' Recently Used Documents Folder
    Private Const CSIDL_SENDTO = 9                     ' Send To Folder
    Private Const CSIDL_STARTMENU = 11                 ' Start Menu Folder
    Private Const CSIDL_DESKTOPDIRECTORY = 16          ' Desktop Folder
    Private Const CSIDL_NETHOOD = 19                   ' Network Neighborhood Folder
    Private Const CSIDL_TEMPLATES = 21                 ' Document Templates Folder
    Private Const CSIDL_COMMON_STARTMENU = 22          ' Common Start Menu Folder
    Private Const CSIDL_COMMON_PROGRAMS = 23           ' Common Program Groups
    Private Const CSIDL_COMMON_STARTUP = 24            ' Common Startup Group Folder
    Private Const CSIDL_COMMON_DESKTOPDIRECTORY = 25   ' Common Desktop Folder
    Private Const CSIDL_APPDATA = 26                   ' Application Data Folder
    Private Const CSIDL_PRINTHOOD = 27                 ' Printers Folder
    Private Const CSIDL_COMMON_FAVORITES = 31          ' Common Favorites Folder
    Private Const CSIDL_INTERNET_CACHE = 32            ' Temp. Internet Files Folder
    Private Const CSIDL_COOKIES = 33                   ' Cookies Folder
    Private Const CSIDL_HISTORY = 34                   ' History Folder



#If VBA7 Then
    Private Declare PtrSafe Function SHGetFolderPath Lib "shell32" ( _
        ByVal hwndOwner As LongPtr, _
        ByVal nFolder As Integer, _
        ByVal hToken As LongPtr, _
        ByVal dwFlags As Long, _
        ByRef pszPath As String) As Long
    Public Declare PtrSafe Function SHGetSpecialFolderLocation _
        Lib "shell32" (ByVal hwnd As LongPtr, _
        ByVal nFolder As Long, ppidl As Long) As Long
    Public Declare PtrSafe Function SHGetPathFromIDList _
        Lib "shell32" Alias "SHGetPathFromIDListA" _
        (ByVal Pidl As Long, ByVal pszPath As String) As Long
    Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pvoid As Long)
#Else
    Private Declare Function SHGetFolderPath Lib "shell32" ( _
        ByVal hwndOwner As Long, _
        ByVal nFolder As Integer, _
        ByVal hToken As Long, _
        ByVal dwFlags As Long, _
        ByRef pszPath As String) As Long
    Public Declare Function SHGetSpecialFolderLocation _
        Lib "shell32" (ByVal hwnd As Long, _
        ByVal nFolder As Long, ppidl As Long) As Long
    Public Declare Function SHGetPathFromIDList _
        Lib "shell32" Alias "SHGetPathFromIDListA" _
        (ByVal Pidl As Long, ByVal pszPath As String) As Long
    Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pvoid As Long)
#End If


Function GetFileNameOnly(ByRef aPath As String, Optional IncludePath As Boolean = False) As String
' Strips the file extension from aPath
    Dim i As Long, PathSep As String
    Dim s As String
    
    PathSep = Application.PathSeparator
    
    s = aPath
    
    If InStr(aPath, ".") > 0 Then
        i = Len(aPath)
        Do While Mid$(aPath, i, 1) <> "."
            i = i - 1
        Loop
        If i > 1 Then s = Left$(aPath, i - 1) Else s = "" ' If there is a period, but it's the first character in aPath, return nothing
    End If
    
    ' Now, strip a potential path, if desired
    If Not IncludePath Then
        If InStr(s, PathSep) > 0 Then
            i = Len(s)
            Do While Mid$(s, i, 1) <> PathSep
                i = i - 1
            Loop
            If i <> Len(s) Then s = Mid$(s, i + 1) Else s = "" ' If there is a path separator, but it's the last character in aPath, return nothing
        End If
    End If
    
    GetFileNameOnly = s

End Function



Function GetFileExtOnly(ByRef aPath As String) As String
' This function will return the file extension in the path at hand including a leading dot
' If no file extension is found, an empty string is returned
    Dim i As Long
    If InStr(aPath, ".") > 0 Then
        i = Len(aPath)
        Do While Mid$(aPath, i, 1) <> "."
            i = i - 1
        Loop
        GetFileExtOnly = Mid$(aPath, i)
    Else ' No dot found, so no extension...
        GetFileExtOnly = ""
    End If
End Function



Function GetPath(ByRef aPath As String) As String
' Returns the path portion of the path/filename/extension string that is handed over.
' Will have a trailing "\". Will be EMPTY if no path separator is found.
    
    Dim i As Long, PathSep As String
    
    If Len(aPath) Then
    
        PathSep = Application.PathSeparator
        
        If InStr(aPath, PathSep) > 0 Then
            i = Len(aPath)
            Do While Mid$(aPath, i, 1) <> PathSep
                i = i - 1
            Loop
            GetPath = Left$(aPath, i)
        Else ' No path separator found, so no path...
            GetPath = ""
        End If
        
    End If

End Function



Function CompletePath(aPath As String) As String

    Dim PathSep As String
    
    PathSep = Application.PathSeparator
    
    If Right$(aPath, 1) = PathSep Then
        CompletePath = aPath
    Else
        CompletePath = aPath & PathSep
    End If

End Function


Function ValidFileNameString(ByRef fn As String) As Boolean

    Dim i As Long
    
    ValidFileNameString = True
    For i = 1 To Len(fn)
        If InStr(InvalidFileNameChars, Mid$(fn, i, 1)) <> 0 Then
            ValidFileNameString = False
            Exit Function
        End If
    Next i
    
End Function


Function MakeValidFileName(ByVal fn As String, Optional UseUnderscores As Boolean = False) As String
' Deletes each 'illegal' character in fn
' Changed declaration for version 2.71: "ByVal fn As String" instead of "ByRef fn As String"
' to avoid inadvertent messing around with the calling procedure's variable.
    
    Dim i As Long
    Dim c As String
    
    If UseUnderscores Then c = "_" Else c = ""
    
    i = 1
    Do Until i > Len(fn)
        If InStr(InvalidFileNameChars & Chr(13) & Chr(10), Mid$(fn, i, 1)) <> 0 Then
            fn = Left$(fn, i - 1) & c & Mid$(fn, i + 1)
        Else
            i = i + 1
        End If
    Loop
    
    MakeValidFileName = fn
End Function



Function CreatePath(ByRef aPath As String) As Boolean
' This function will try to generate the aPath
' in the file system. Unlike MkDir, it can create
' several subfolders at once.
' If successful, it returns TRUE.
' Note that the flag "vbHidden" has to be included
' in the Dir() calls so that hidden folders are
' recognized.

    Dim i As Long
    Dim SubPath As String
    Dim PathSep As String
    
    PathSep = Application.PathSeparator
    aPath = AddPathSep(aPath)
    i = InStr(aPath, PathSep)
    Do
        SubPath = Left$(aPath, i)
        ' Try to make the sub directory (it may exist already)
        On Error Resume Next
            MkDir SubPath
            
            ' Fix for version 2.71: error report 4KA3VR
            ' Do not crash ugly if creating the path failed;
            ' move the If(Len(Dir... statement inside the
            ' On Error Resume Next bracket.
            
            ' If the sub directory still does not exist,
            ' creating the path failed -- exit the function
            If Len(Dir(RemovePathSep(SubPath), vbDirectory Or vbHidden Or vbSystem)) = 0 Then GoTo ErrorExit
        On Error GoTo 0
        i = InStr(i + 1, aPath, PathSep)
    Loop Until i = 0 ' Loop until no more path separators are found
        
    CreatePath = Len(Dir(RemovePathSep(aPath), vbDirectory Or vbHidden Or vbSystem)) <> 0
ErrorExit:
End Function


Function ChDirEx(path As String) As String
' Changes the working directory and drive; returns the old working directory
    ChDirEx = CurDir
    If path Like "?:*" Then ChDrive Left$(path, 2)
    ChDir path
End Function


Private Function SpecFolder(ByVal lngFolder As Long) As String
' http://msdn.microsoft.com/en-us/library/aa140088(office.10).aspx
    Dim lngPidlFound As Long
    Dim lngFolderFound As Long
    Dim lngPidl As Long
    Dim strPath As String

    strPath = Space(MAX_PATH)
    lngPidlFound = SHGetSpecialFolderLocation(0, lngFolder, lngPidl)
    If lngPidlFound = 0 Then
        lngFolderFound = SHGetPathFromIDList(lngPidl, strPath)
        If lngFolderFound Then
            SpecFolder = Left$(strPath, _
                InStr(1, strPath, vbNullChar) - 1)
        End If
    End If
    CoTaskMemFree lngPidl
End Function


Function GetSpecialFolderDesktop() As String
    GetSpecialFolderDesktop = SpecFolder(CSIDL_DESKTOPDIRECTORY)
End Function


Function GetSpecialFolderMyDocuments() As String
    GetSpecialFolderMyDocuments = SpecFolder(CSIDL_PERSONAL)
End Function

Function FileExists(filename As String) As Boolean
    On Error Resume Next
    FileExists = Len(Dir(filename, vbArchive Or vbDirectory Or vbHidden Or vbSystem)) <> 0
End Function

Function AddPathSep(s As String) As String
' This function adds a path separator to the
' path contained in s, unless there already
' is one.

    Dim NewPath As String
    Dim PathSep As String
    PathSep = Application.PathSeparator
    NewPath = Trim$(s)
    If (Len(NewPath) > 0) And (Right$(NewPath, 1) <> PathSep) Then
        NewPath = NewPath & PathSep
    End If
    AddPathSep = NewPath
    
End Function

Function RemovePathSep(s As String) As String

    Dim NewPath As String
    Dim PathSep As String
    
    PathSep = Application.PathSeparator
    NewPath = Trim$(s)
    If (Right$(NewPath, 1) = PathSep) Then
        NewPath = Left$(NewPath, Len(NewPath) - 1)
        
        ' Make sure not to remove the path separator
        ' after a drive symbol (root path)
        If Right$(NewPath, 1) = ":" Then NewPath = NewPath & PathSep
    End If
    RemovePathSep = NewPath

End Function

