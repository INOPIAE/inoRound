VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGetTextObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' == XL TOOLBOX   (c) 2008-2013 Daniel Kraus    Licensed under GPLv2
' ==================================================================
' == DanielsXLToolbox.clsGetTextObject
' ==
' == Created: 05-Nov-11 17:47
' ==
' == Stores information about a gettext translatable string (msgid).
' == See: www.gnu.org/s/hello/manual/gettext/PO-files.html


Option Explicit

    Private mMsgId As String
    Private mHash As Long
    Private mTranslatorComments As String
    Private mExtractedComments As String
    Private mContext As String
    Private mReferences As String
    Private mFlags As String
    Private mPreviousMsgId As String
    Private mMsgStr As String
    
    Private mStatus As Long
    Private mStatusMsg As String
    
    Private Const EOT = 4

Property Get msgId() As String
    msgId = mMsgId
End Property

Property Let msgId(id As String)
    mMsgId = id
    ComputeHashForMsgId
End Property

Property Get hash() As Long
    hash = mHash
End Property

Property Get translatorComments() As String
    translatorComments = mTranslatorComments
End Property

Property Get extractedComments() As String
    extractedComments = mExtractedComments
End Property

Property Get references() As String
    references = mReferences
End Property

Property Get flags() As String
    flags = mFlags
End Property

Property Get previousMsgId() As String
    previousMsgId = mPreviousMsgId
End Property

Property Get context() As String
    context = mContext
End Property

Property Let context(str As String)
    mContext = context
    ComputeHashForMsgId
End Property

Property Get msgStr() As String
    msgStr = mMsgStr
End Property

Property Get translation() As String
' Returns the translated string (i.e., MsgStr), if one exists.
' Otherwise, returns the untranslated string (MsgId).
    If Len(mMsgStr) Then
        translation = mMsgStr
    Else
        translation = mMsgId
    End If
End Property

Property Get status() As Long
    status = mStatus
End Property

Property Get statusMsg() As String
    statusMsg = mStatusMsg
End Property


Function ReadFromFile(file As clsTextFile) As Boolean
' Reads a PO object from the file represented by FileHandle.
' FileHandle must refer to an open text file.
' Returns True if successful, False if it failed.
    Dim line As String
    Dim c As String * 2
    
    On Error GoTo ErrorHandle
    
    ' Read lines from the file until we hit an empty line
    Do
        line = Trim$(file.GetLine)
        If Left$(line, 1) = "#" Then
            c = Mid$(line, 2, 2)
            If c = "  " Then
                mTranslatorComments = Mid$(line, 4)
            ElseIf c = ". " Then
                mExtractedComments = Mid$(line, 4)
            ElseIf c = ": " Then
                mReferences = Mid$(line, 4)
            ElseIf c = ", " Then
                mFlags = Mid$(line, 4)
            ElseIf c = "| " Then
                mPreviousMsgId = Mid$(line, 4)
            End If
        ElseIf Left$(line, 8) = "msgctxt " Then
            mContext = TrimQuotes(Mid$(line, 9))
        ElseIf Left$(line, 6) = "msgid " Then
            mMsgId = ReadMultiLine(file, Mid$(line, 7))
            ' mMsgId = UnescapeQuotes(TrimQuotes((Mid$(Line, 7))))
        ElseIf Left$(line, 7) = "msgstr " Then
            mMsgStr = ReadMultiLine(file, Mid$(line, 8))
            ' mMsgStr = UnescapeQuotes(TrimQuotes(Mid$(Line, 8)))
        End If
    Loop Until Len(line) = 0
            
    ComputeHashForMsgId
    ReadFromFile = True

ErrorHandle:
    mStatus = Err.Number
    If mStatus Then
        mStatusMsg = Err.Description
        Debug.Print "clsGetTextObject:ReadFromFile", mStatusMsg
    Else
        mStatusMsg = ""
    End If
End Function


Function WriteToFile(f As clsTextFile) As Boolean
' Writes the object to the file represented by FileHandle,
' which must be open for write access.
    On Error GoTo ErrorHandle
    
    With f
        .PutLine "" ' Output whitespace as a delimiter
        If Len(mTranslatorComments) Then .PutLine "#  " & mTranslatorComments
        If Len(mExtractedComments) Then .PutLine "#. " & mExtractedComments
        If Len(mReferences) Then .PutLine "#: " & mReferences
        If Len(mFlags) Then .PutLine "#, " & mFlags
        If Len(mPreviousMsgId) Then .PutLine "#|  " & mPreviousMsgId
        If Len(mContext) Then .PutLine "msgctxt """ & mContext & """"
        WriteMultiLine f, "msgid ", mMsgId
        WriteMultiLine f, "msgstr ", mMsgStr
    End With ' f
    
    WriteToFile = True
    
ErrorHandle:
    mStatus = Err.Number
    If mStatus Then
        mStatusMsg = Err.Description
    Else
        mStatusMsg = ""
    End If
End Function

Function AddReference(module As String, line As String) As Boolean
' Adds the reference to the object.
' Returns True if the reference is new.
    Dim ref As String
    ref = BuildReference(module, line)
    If InStr(mReferences, ref) = 0 Then
        mReferences = mReferences & " " & ref
        AddReference = True
    End If
End Function

Sub Initialize(msgId As String, Optional msgStr As String, Optional module As String, _
    Optional line As String, Optional context As String)
' Initializes the object with arguments
    mMsgId = msgId
    If Len(msgStr) Then mMsgStr = msgStr
    If Len(Trim$(module)) Then mReferences = BuildReference(module, line)
    If Len(Trim$(context)) Then mContext = context
    ComputeHashForMsgId
End Sub

' ========================================================
' == Private class methods
' ========================================================

Private Function TrimQuotes(s As String) As String
' Removes one pair of leading and trailing quotes, if one is present
    Dim str As String
    str = Trim$(s) ' Remove whitespace, if any
    If (Left$(str, 1) = """" And Right$(str, 1) = """") Or _
            (Left$(str, 1) = "'" And Right$(str, 1) = "'") Then
        str = Mid$(str, 2, Len(str) - 2)
    End If
    TrimQuotes = str
End Function

Private Sub ComputeHashForMsgId()
' Computes that hash value for the current MsgId
' This is a wrapper for the xltoolbox dll function compute_hashval(),
' which in turn is taken directly from the hash.h source code of
' the GetText library (GNU GPL v. 2).
    If Len(mContext) Then
        mHash = ComputeHash(mContext & Chr(EOT) & mMsgId)
    Else
        mHash = ComputeHash(mMsgId)
    End If
End Sub

Private Function BuildReference(module As String, line As String) As String
' Builds the reference string using the Module and Line components.
    If Len(Trim$(line)) Then
        BuildReference = Trim$(module) & ":" & Trim$(line)
    Else
        BuildReference = Trim$(module)
    End If
End Function

Private Function EscapeQuotes(s As String) As String
' "Escapes" quotation marks by adding a backslash.
    Dim i As Long
    i = InStr(s, """")
    While i
        s = Left$(s, i - 1) & "\" & Mid$(s, i)
        i = InStr(i + 2, s, """")
    Wend
    
    ' Also escape the occasional line break...
    i = InStr(s, vbNewLine)
    While i
        s = Left$(s, i - 1) & "\n" & Mid$(s, i + Len(vbNewLine))
        i = InStr(i + 2, s, vbNewLine)
    Wend
    
    EscapeQuotes = s
End Function

Private Function UnescapeQuotes(s As String) As String
' Reverse of the above EscapeQuotes function
    Dim i As Long
    i = InStr(s, "\""")
    While i
        s = Left$(s, i - 1) & Mid$(s, i + 1)
        i = InStr(s, "\""")
    Wend
    
    UnescapeQuotes = s
End Function

Private Function ReadMultiLine(f As clsTextFile, FirstLine As String) As String
' Reads one or several lines from a PO(T) file. Lines may be continuated
' like this:
'     msgid "This is a very "
'     "long line."
' NB: The FirstLine parameter must have the initial "msgid " (or whatever)
' stripped from it when calling this function.
    Dim line As String
    Dim s As String
    Dim multi As String
    
    multi = UnescapeQuotes(TrimQuotes(FirstLine))
    Do
        line = f.GetLine()
        s = Trim$(line)
        While Left$(s, 1) = vbTab
            s = Mid$(s, 2)
        Wend
        If Left$(s, 1) = """" Then
            s = UnescapeQuotes(TrimQuotes(Trim$(line)))
            multi = multi & s
        Else
            f.PutLine line
            Exit Do
        End If
    Loop
    ReadMultiLine = multi
End Function

Private Sub WriteMultiLine(f As clsTextFile, FirstLinePrefix As String, line As String)
' Writes the Line in a PO file-style manner to the file.
' The string will be split at line breaks ("\n") to produce more
' beautiful output. If there are no line breaks, the entire
' string is written into one line
    If InStr(line, "\n") Then
        Dim s() As String
        Dim i As Long
        With f
            s = Split(EscapeQuotes(line), "\n")
            ' For clarity, the first line of a multi-line string will have an
            ' empty string (see GNU gettext documentation)
            .PutLine FirstLinePrefix & """"""
            For i = 0 To UBound(s) - 1
                .PutLine """" & s(i) & "\n"""
            Next i
            .PutLine """" & s(i) & """"
        End With ' f
    Else
        ' No newline characters found
        f.PutLine FirstLinePrefix & """" & EscapeQuotes(line) & """"
    End If
End Sub
