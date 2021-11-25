VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGetTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' == XL TOOLBOX   (c) 2008-2013 Daniel Kraus    Licensed under GPLv2
' ==================================================================
' == DanielsXLToolbox.clsGetTextFile
' ==
' == Created: 05-Nov-11 17:43
' ==
' == Provides an 'API' to a GNU gettext PO/MO file.


Option Explicit
    
    Private Const EOT = 4

    Enum GetTextStatus
        GetTextStatusOK = 0
        GetTextStatusInvalidMoFile = -100
    End Enum
    
    Private mStatus As GetTextStatus
    Private mStatusMsg As String
    
    Private mFileName As String
    Private mFile As clsTextFile
    Private mUpdated As Boolean
    
    ' The PO objects are held in a VBA collection which uses GNU GetText
    ' hashes as keys. Since the VBA collection works with string keys
    ' only, we need to convert the hash values to strings... But still,
    ' this should be more efficient than using the MsgId strings
    ' themselves as keys, as the converted values are much shorter strings
    ' and are thus more efficient to compare.
    ' \todo Increase efficiency by implementing a collection that uses true integer keys.
    Private mObjects As Collection
    Private mHeaderLines() As String
    Private mNumHeaderLines As Long
    
    ' PO file metadata
    Private mProjectId  As String
    Private mReportBugsTo As String
    Private mPotCreationDate As String
    Private mPoRevisionDate As String
    Private mLastTranslator As String
    Private mLanguageTeam As String
    Private mLanguage As String
    Private mMimeVersion As String
    Private mContentType As String
    Private mContentEncoding As String
    Private mPluralForms As String
    
    
' To be able to use MO files with UTF-8 encoding, we need some Win32 API functions
' that convert UTF-8 to UTF-16. VBA encodes all strings in 2-byte-wide UTF-16 internally.

#If VBA7 Then
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, _
        ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
        
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, _
        ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, _
        ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As LongPtr) As Long
        
    Private Declare PtrSafe Function ReverseByteOrderLong Lib "xltoolbox.dll" (ByVal n As Long) As Long
#Else
    Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, _
        ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    
    Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    
    Private Declare Function ReverseByteOrderLong Lib "xltoolbox.dll" (ByVal n As Long) As Long
#End If
                         
Private Const CP_UTF8 As Long = 65001

    
    

Function ReadFile(filename As String) As Boolean
' Reads a PO/MO file
        
    If UCase$(Right$(filename, 3)) = ".PO" Then
        ReadFile = ReadPoFile(filename)
    ElseIf UCase$(Right$(filename, 3)) = ".MO" Then
        ReadFile = ReadMoFile(filename)
    End If
    
End Function


Function WriteFile(filename As String) As Boolean
    On Error GoTo ErrorHandle
    
    Dim o As clsGetTextObject
    Dim fail As Boolean
    
    If mUpdated Then mPoRevisionDate = GetISO8601TimeStamp
    
    Set mFile = New clsTextFile
    If Not mFile.OpenToWrite(filename) Then
        Err.Raise 9999, "clsGetTextFile.WriteFile", "Failed to open file for writing (clsTextFile.OpenToWrite)."
    End If
    
    If Not WriteHeader(mFile) Then GoTo ErrorHandle
    
    For Each o In mObjects
        If Not o.WriteToFile(mFile) Then
            fail = True
            Exit For
        End If
    Next o
    
    If Not fail Then fail = Not mFile.CloseFile
    WriteFile = Not fail
    If Not fail Then mUpdated = False
    
ErrorHandle:
    mStatus = Err.Number
    If mStatus Then
        mStatusMsg = Err.Description
    Else
        mStatusMsg = ""
    End If

End Function

Function GetText(msgId As String) As String
' Looks up MsgId in the PO file objects
    Dim h As Long
    Dim o As clsGetTextObject
    
    h = ComputeHash(msgId)
    On Error Resume Next
        Set o = mObjects(CStr(h))
    On Error GoTo 0
    If Not (o Is Nothing) Then
        ' If there is a translation, return it; otherwise, return the
        ' untranslated string.
        GetText = o.translation
    Else
        ' If the MsgId hash was not found in the collection,
        ' return the original MsgId
        GetText = msgId
    End If
End Function

Function GetTextContext(context As String, msgId As String) As String
    Dim h As Long
    Dim o As clsGetTextObject
    
    If Len(context) Then
        h = ComputeHash(context & Chr(EOT) & msgId)
    Else
        h = ComputeHash(msgId)
    End If
    
    On Error Resume Next
        Set o = mObjects(CStr(h))
    On Error GoTo 0
    If Not (o Is Nothing) Then
        ' If there is a translation, return it; otherwise, return the
        ' untranslated string.
        GetTextContext = o.translation
    Else
        ' If the MsgId hash was not found in the collection,
        ' return the original MsgId
        GetTextContext = msgId
    End If
End Function

Function AddMsgId(msgId As String, module As String, line As String, Optional context As String) As Boolean
' Adds (or updates) a MsgId to the file.
' If MsgId exists already, the location ("<Module>:<Line>") is added to the
' reference comment of the PO object.
' If MsgId is from a Form, 'Module' should be the form's name, and 'Line'
' the name of the Form Control.
' Returns True if the MsgId is new.
' \todo: Ability to mark MsgIds as 'fuzzy'.
    Dim h As Long
    Dim o As clsGetTextObject
    
    h = ComputeHash(msgId)
    On Error Resume Next
        Set o = Nothing
        Set o = mObjects(CStr(h))
    On Error GoTo 0
    
    If Not (o Is Nothing) Then
        o.AddReference module, line
    Else
        Set o = New clsGetTextObject
        o.Initialize msgId, module:=module, line:=line, context:=context
        mObjects.Add o, CStr(h)
        AddMsgId = True
    End If
    mUpdated = True
End Function

Sub AddHeaderLine(line As String)
' Adds a comment line to the file header.
    Dim s As String
    s = line
    If Left$(s, 1) <> "#" Then s = "#" & s
    mNumHeaderLines = mNumHeaderLines + 1
    If mNumHeaderLines > UBound(mHeaderLines) Then
        ReDim Preserve mHeaderLines(UBound(mHeaderLines) + 5)
    End If
    mHeaderLines(mNumHeaderLines) = s
End Sub

' =================================================================
' Class properties
' =================================================================

Property Get status() As Long
    status = mStatus
End Property

Property Get statusMsg() As String
    statusMsg = mStatusMsg
End Property

Property Get MetaProjectIdVersion() As String
    MetaProjectIdVersion = mProjectId
End Property

Property Let MetaProjectIdVersion(str As String)
    mProjectId = str
End Property

Property Get MetaReportMsgidBugsTo() As String
    MetaReportMsgidBugsTo = mReportBugsTo
End Property

Property Let MetaReportMsgidBugsTo(str As String)
    mReportBugsTo = str
End Property

Property Get MetaPotCreationDate() As String
    MetaPotCreationDate = mPotCreationDate
End Property

Property Get MetaPoRevisionDate() As String
    MetaPoRevisionDate = mPoRevisionDate
End Property

Property Get MetaLastTranslator() As String
    MetaLastTranslator = mLastTranslator
End Property

Property Let MetaLastTranslator(str As String)
    mLastTranslator = str
End Property

Property Get MetaLanguageTeam() As String
    MetaLanguageTeam = mLanguageTeam
End Property

Property Let MetaLanguageTeam(str As String)
    mLanguageTeam = str
End Property

Property Get MetaLanguage() As String
    MetaLanguage = mLanguage
End Property

Property Let MetaLanguage(str As String)
    mLanguage = str
End Property

Property Get MetaMimeVersion() As String
    MetaMimeVersion = mMimeVersion
End Property

Property Let MetaMimeVersion(str As String)
    mMimeVersion = str
End Property

Property Get MetaContentType() As String
    MetaContentType = mContentType
End Property

Property Let MetaContentType(str As String)
    mContentType = str
End Property

Property Get MetaContentEncoding() As String
    MetaContentEncoding = mContentEncoding
End Property

Property Let MetaContentEncoding(str As String)
    mContentEncoding = str
End Property

Property Get MetaPluralForms() As String
    MetaPluralForms = mPluralForms
End Property

Property Let MetaPluralForms(str As String)
    mPluralForms = str
End Property

Property Get Saved() As Boolean
    Saved = Not mUpdated
End Property


' =================================================================
' Private methods of the class
' =================================================================

Private Function ReadPoFile(filename As String) As Boolean
    Dim o As clsGetTextObject
    Set mFile = New clsTextFile
    
    If mFile.OpenToRead(filename) Then
        If ReadHeader Then
            While Not mFile.EOF
                Set o = New clsGetTextObject
                o.ReadFromFile mFile
                If o.status = 0 Then mObjects.Add o, CStr(o.hash)
            Wend
            ReadPoFile = True
        End If
    End If

ErrorHandle:
' There used to be an "On error goto ErrorHandle" statement at the
' beginning of this function, but somehow it would cause the ReadFile = True
' statement to never be executed??? Even though the file was read alright?

    mFile.CloseFile ' Important! Otherwise, we cannot write to this file.
    mStatus = Err.Number
    If mStatus Then
        Stop
        mStatusMsg = Err.Description
    Else
        mStatusMsg = ""
    End If
End Function

Private Function ReadHeader() As Boolean
' Reads the PO file header from the file
    
    On Error GoTo ErrorHandle
    Dim line As String
    Dim metaName As String
    Dim metaContent As String
    
    Do
        line = mFile.GetLine
        If (Left$(line, 2) = "# ") Or (Len(Trim$(line)) = 0) Then
            mNumHeaderLines = mNumHeaderLines + 1
            If mNumHeaderLines > UBound(mHeaderLines) Then
                ReDim Preserve mHeaderLines(1 To UBound(mHeaderLines) + 10)
            End If
            mHeaderLines(mNumHeaderLines) = line
        End If
    Loop Until ((Left$(line, 2) <> "# ") And Len(Trim$(line)) > 0) Or mFile.EOF
    
    ' Read the metadata
    With mFile
        .PutLine line
        If Not mFile.EOF Then
            line = .GetLine(Trimmed:=True)
            If line = "msgid """"" Then
                line = .GetLine(Trimmed:=True)
                If line = "msgstr """"" Then
                    While ReadMetadataLine(line, metaName, metaContent)
                        Select Case metaName
                            Case "Project-Id-Version": mProjectId = metaContent
                            Case "Report-Msgid-Bugs-To": mReportBugsTo = metaContent
                            Case "POT-Creation-Date": mPotCreationDate = metaContent
                            Case "PO-Revision-Date": mPoRevisionDate = metaContent
                            Case "Last-Translator": mLastTranslator = metaContent
                            Case "Language-Team": mLanguageTeam = metaContent
                            Case "Language": mLanguage = metaContent
                            Case "MIME-Version": mMimeVersion = metaContent
                            Case "Content-Type": mContentType = metaContent
                            Case "Content-Transfer-Encoding": mContentEncoding = metaContent
                            Case "Plural-Forms": mPluralForms = metaContent
                        End Select
                    Wend
                    ' Read on until the first blank line
                    While (Len(line) > 0) And (Not .EOF)
                        line = .GetLine(Trimmed:=True)
                    Wend
                End If
            End If
        End If
    End With ' mFile
    
    
    ' The line below causes runtime errors, since files opened for
    ' input access read cannot handle Seek commands. Therefore, a
    ' mini file buffer was implemented.
    ' If Not EOF(filehandle) Then Seek #filehandle, Loc(filehandle) - 1
    If Len(line) Then mFile.PutLine line
    
    ReadHeader = True
ErrorHandle:
End Function

Private Function ReadMetadataLine(ByRef line As String, ByRef metaName As String, ByRef metaContent As String) As Boolean
' Reads a metadata line from mFile, strips the quotes surrounding it,
' and splits it by the colon.
' Returns True if the format is syntactically correct.
' Returns the name for the meta information as MetaName.
' Returns the content of the meta information as MetaContent
    Dim i As String
    line = mFile.GetLine(Trimmed:=True)
    If (Left$(line, 1) = """") And (Right$(line, 1) = """") Then
        line = Mid$(line, 2, Len(line) - 2)
        i = InStr(line, ":")
        If i Then
            metaName = Trim$(Left$(line, i - 1))
            ' The content is trimmed only on the left side.
            metaContent = LTrim$(Mid$(line, i + 1))
            While Right$(line, 2) = "\n"
                line = Left$(line, Len(line) - 2)
            Wend
            ReadMetadataLine = True
        End If
    End If
End Function

Private Function WriteHeader(f As clsTextFile) As Boolean
' Writes the PO file header to the text file f
    On Error GoTo ErrorHandle
    Dim i As Long
    
    ' Write the leading comment
    With f
        For i = 1 To mNumHeaderLines
            .PutLine mHeaderLines(i)
        Next i
        
        ' If there are header comments, they need to be separated
        ' from the rest of the file by a blank line.
        If mNumHeaderLines Then
            If Len(Trim$(mHeaderLines(mNumHeaderLines))) Then .PutLine " "
        End If
        
        ' Write the metadata
        .PutLine "msgid """""
        .PutLine "msgstr """""
        .PutLine """Project-Id-Version: " & mProjectId & "\n"""
        .PutLine """Report-Msgid-Bugs-To: " & mReportBugsTo & "\n"""
        .PutLine """POT-Creation-Date: " & mPotCreationDate & "\n"""
        .PutLine """PO-Revision-Date: " & mPoRevisionDate & "\n"""
        .PutLine """Last-Translator: " & mLastTranslator & "\n"""
        .PutLine """Language-Team: " & mLanguageTeam & "\n"""
        .PutLine """Language: " & mLanguage & "\n"""
        .PutLine """MIME-Version: " & mMimeVersion & "\n"""
        .PutLine """Content-Type: " & mContentType & "\n"""
        .PutLine """Content-Transfer-Encoding: " & mContentEncoding & "\n"""
        .PutLine """Plural-Forms: " & mPluralForms & "\n"""
    End With ' f
    WriteHeader = True
ErrorHandle:
    If Err Then
        mStatus = Err.Number
        mStatusMsg = Err.Description
        Debug.Print mStatus, mStatusMsg
    End If
End Function


Private Function ReadMoFile(filename As String) As Boolean
' Reads a GNU MO file.
' Note: This function does not support all MO file features yet,
' e.g. pluralization or contexts.
' A possible hash table at the end of the file is ignored.
    Dim f As Long
    Dim bigEnd As Boolean
    Dim i As Long
    Dim j As Long
    Dim nStrings As Long
    Dim msgIdLength() As Long
    Dim msgIdOffset() As Long
    Dim msgStrLength() As Long
    Dim msgStrOffset() As Long
    Dim bytes() As Byte
    Dim msgId As String
    Dim msgStr As String
    Dim msgCtxt As String
    Dim o As clsGetTextObject
    
    On Error GoTo ErrorExit
    f = FreeFile
    Open filename For Binary Access Read As #f
    
    Get #f, , i
    
    ' Check for magic number (in different endiannesses)
    If (i = &H950412DE) Or (i = &HDE120495) Then
        ' If the file has big-endian byte order, we will need
        ' to convert the Long integers...
        bigEnd = (i = &HDE120495)
    
        Get #f, , i ' File format revision
        If i = 0 Then
            nStrings = GetLongFromFile(f, bigEnd)
            
            ReDim msgIdOffset(0 To nStrings) As Long
            ReDim msgIdLength(1 To nStrings) As Long
            ReDim msgStrOffset(0 To nStrings) As Long
            ReDim msgStrLength(1 To nStrings) As Long
            
            msgIdOffset(0) = GetLongFromFile(f, bigEnd)
            msgStrOffset(0) = GetLongFromFile(f, bigEnd)
            
            Seek #f, 1 + msgIdOffset(0) ' Seek works 1-based, offsets are zero-based
            For i = 1 To nStrings
                msgIdLength(i) = GetLongFromFile(f, bigEnd)
                msgIdOffset(i) = GetLongFromFile(f, bigEnd)
            Next i
            
            Seek #f, 1 + msgStrOffset(0)
            For i = 1 To nStrings
                msgStrLength(i) = GetLongFromFile(f, bigEnd)
                msgStrOffset(i) = GetLongFromFile(f, bigEnd)
            Next i
            
            For i = 1 To nStrings
                If msgIdLength(i) Then
                    ReDim bytes(1 To msgIdLength(i))
                    Get #f, 1 + msgIdOffset(i), bytes
                    msgId = Utf8ToStr(bytes)
                    
                    ' Since the msgId may contain the context information
                    ' in the form <context><EOT><msgId>, we now check for the
                    ' EOT byte (ASCII code 4) and separate the two.
                    j = InStr(msgId, Chr(EOT))
                    If j Then
                        msgCtxt = Left$(msgId, j - 1)
                        msgId = Mid$(msgId, j + 1)
                    End If
                Else
                    msgId = vbNullString
                End If
                
                If msgStrLength(i) Then
                    ReDim bytes(1 To msgStrLength(i))
                    Get #f, 1 + msgStrOffset(i), bytes
                    msgStr = Utf8ToStr(bytes)
                Else
                    msgStr = vbNullString
                End If
                
                
                ' The strings in the MO file have newlines encoded
                ' as Chr(10). We need to escape those to "\n" to be
                ' able to compare the hash values with the strings
                ' that are hard-coded in the add-in; those have "\n"
                ' in them.
                msgId = EscapeStr(msgId)
                msgStr = EscapeStr(msgStr)
                
                Set o = New clsGetTextObject
                o.Initialize msgId, msgStr, context:=msgCtxt
                mObjects.Add o, CStr(o.hash)
                ' Debug.Print o.hash, msgCtxt, msgId
            Next i
            
            ReadMoFile = True
        Else
            mStatus = GetTextStatusInvalidMoFile
            mStatusMsg = "Invalid MO file version (must be 0)."
        End If
    Else
        mStatus = GetTextStatusInvalidMoFile
        mStatusMsg = "The file does not contain the correct MO file magic numbers."
    End If
    
ErrorExit:
    If Err Then
        mStatus = Err.Number
        mStatusMsg = Err.Description
    End If
End Function

Private Sub Class_Initialize()
    ' Set mObjects = CreateObject("Scripting.Dictionary")
    Set mObjects = New Collection
    ReDim mHeaderLines(1 To 10)
    
    mPluralForms = "nplurals=2; plural=(n!=1);"
    mProjectId = "0.0"
    mReportBugsTo = "EMAIL-ADDRESS"
    mPotCreationDate = GetISO8601TimeStamp
    mPoRevisionDate = "1970-01-01 12:00:00 +0000"
    mLastTranslator = "LAST TRANSLATOR"
    mLanguageTeam = "TEAM NAME"
    mLanguage = "EN-US"
    mContentEncoding = "8bits"
    mContentType = "text/plain; charset=UTF-8" ' \todo: Check if UTF-8 is true
    mMimeVersion = "1.0"
End Sub



Private Function GetISO8601TimeStamp(Optional IncludeTZOffset As Boolean = True) As String
' This function returns the current date/time
' formatted according to ISO 8601 (without timezone info),
' i.e. 2009-07-02 20:02 is July 2, 2009, 8:02 pm.
' Note that we cannot separate the date from the time
' by a "T", because that would make the string
' uninterpretable for Excel, and we would have to
' create our own ISO8601-to-Date conversion function.
    
    If IncludeTZOffset Then
        GetISO8601TimeStamp = Format(Now, "yyyy-MM-dd hh:mm:ss") & " " & TimeZoneOffsetStr
    Else
        GetISO8601TimeStamp = Format(Now, "yyyy-MM-dd hh:mm:ss")
    End If
End Function

Private Function StrToUtf8(ByRef s As String) As Byte()
    Dim i As Long
    #If VBA7 Then
        Dim ptr As LongPtr
    #Else
        Dim ptr As Long
    #End If
    Dim n As Long
    Dim b() As Byte
    
    i = Len(s)
    ptr = StrPtr(s)
    
    ' Find out how large our buffer has to be
    n = WideCharToMultiByte(CP_UTF8, 0, ptr, i, 0, 0, 0, 0)
    ReDim b(0 To n - 1)
    
    ' Do the actual conversion
    WideCharToMultiByte CP_UTF8, 0, ptr, i, VarPtr(b(0)), n, 0, 0
    StrToUtf8 = b
End Function

Private Function Utf8ToStr(ByRef b() As Byte) As String
    #If VBA7 Then
        Dim ptr As LongPtr
    #Else
        Dim ptr As Long
    #End If
    Dim i As Long
    Dim n As Long
    Dim s As String
    
    n = UBound(b) - LBound(b) + 1 ' The size of the buffer in bytes
    ptr = VarPtr(b(LBound(b))) ' A pointer to the first byte in the buffer
    
    ' Find out how long the string will be and prepare the string
    i = MultiByteToWideChar(CP_UTF8, 0, ptr, n, 0, 0)
    s = Space$(i)
    
    ' Perform the actual conversion
    MultiByteToWideChar CP_UTF8, 0, ptr, n, StrPtr(s), i
    Utf8ToStr = s
    
End Function

Private Function GetLongFromFile(file As Long, bigEndian As Boolean) As Long
' Reads a 32-bit long integer from a binary file and reverses its
' byte order if the file is written in big-endian mode.
    Dim n As Long
    Get #file, , n
    If bigEndian Then n = ReverseByteOrderLong(n)
    GetLongFromFile = n
End Function

