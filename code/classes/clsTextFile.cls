VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==================================================================
' == XL TOOLBOX   (c) 2008-2013 Daniel Kraus    Licensed under GPLv2
' ==================================================================
' == XLDevTools.clsTextFile
' ==
' == Created: 18-Nov-11 14:29
' ==
' == A very simple text file class that enables pushing back text
' == lines into the file. The data is written in UTF-8 encoding.
' == Tipps for how to do this were taken from stackoverflow:
' == http://stackoverflow.com/questions/2524703 (general principle)
' == http://stackoverflow.com/questions/4143524 (stripping the BOM)


Option Explicit

    Private Const CHARSET = "UTF-8"
    
    Private mOpen As Boolean
    Private mToWrite As Boolean
    Private mFileName As String
    Private mLineBuffer() As String
    Private mBufferSize As Long
    Private Const BUFFER_INCREMENT = 5
    
    Private mAdodbStream As Object
    
    ' A couple of ADODB stream constants, so that we do not
    ' need to add the reference to the
    ' "Microsoft ActiveX Data Objects 2.5 Object Library"
    ' to the project.
    ' For an ADODB reference, see:
    ' http://www.w3schools.com/ADO/ado_ref_stream.asp
    Private Const adCRLF = -1
    Private Const adLF = 10
    Private Const adCR = 13
    
    Private Const adTypeBinary = 1
    Private Const adTypeText = 2
    
    Private Const adModeRead = 1
    Private Const adModeWrite = 2
    Private Const adModeReadWrite = 3
    
    Private Const adStateClosed = 0
    Private Const adStateOpen = 1
    Private Const adStateConnecting = 2
    Private Const adStateExecuting = 4
    Private Const adStateFetching = 8
    
    Private Const adWriteChar = 0
    Private Const adWriteLine = 1
    
    Private Const adSaveCreateNotExist = 1
    Private Const adSaveCreateOverWrite = 2

    Private Const adReadAll = -1
    Private Const adReadLine = -2

Property Get EOF() As Boolean
    If mOpen Then
        EOF = (mBufferSize = 0) And mAdodbStream.EOS
    Else
        Err.Raise 9999, "clsTextFile.EOF", "File not open."
    End If
End Property

Function OpenToRead(filename As String) As Boolean
    On Error GoTo ErrorHandle
    If Not mOpen Then
        Set mAdodbStream = CreateObject("adodb.stream")
        With mAdodbStream
            .Open
            .LineSeparator = adCRLF ' Windows-style line endings
            .CHARSET = CHARSET
            .LoadFromFile filename
            ' .Type = adTypeText ' This would cause a run-time errors here
            
            ' Did we succeed?
            If .State = adStateOpen Then
                mFileName = filename
                mOpen = True
                mToWrite = False
                OpenToRead = True
            End If
        End With
    End If
ErrorHandle:
    If Err Then Debug.Print Err.Number, Err.Description
End Function

Function OpenToWrite(filename As String) As Boolean
    On Error GoTo ErrorHandle
    If Not mOpen Then
        Set mAdodbStream = CreateObject("adodb.stream")
        With mAdodbStream
            .Type = adTypeText
            
            ' Important to set the mode to read/write, since we need to be able to read
            ' the stream's data and copy it into a binary stream, in order to save it
            ' without BOM (see Close method).
            .mode = adModeReadWrite
            .LineSeparator = adCRLF ' Windows-style line endings
            .CHARSET = CHARSET
            .Open
            
            ' Did we succeed?
            If .State = adStateOpen Then
                mFileName = filename
                mOpen = True
                mToWrite = True
                OpenToWrite = True
            End If
        End With
    Else
        Err.Raise 9999, "clsTextFile", "File already open."
    End If
ErrorHandle:
    If Err Then Debug.Print Err.Number, Err.Description
End Function

Function CloseFile() As Boolean
' Closes the file. If thefile was opened for writing, this
' will perform the actual write operation. A byte-order mark (BOM)
' is stripped using the technique described at stackoverflow.com
' (question 4143524).
    On Error Resume Next
    If mOpen Then
        If mToWrite Then
            mAdodbStream.Position = 3 ' Skip BOM
            Dim binStream As Object
            Set binStream = CreateObject("adodb.stream")
            binStream.Type = adTypeBinary
            binStream.mode = adModeReadWrite
            binStream.Open
            mAdodbStream.CopyTo binStream
            mAdodbStream.Flush
            binStream.SaveToFile mFileName, adSaveCreateOverWrite
            binStream.Flush
            binStream.Close
            Set binStream = Nothing
        End If
        mAdodbStream.Close
        Set mAdodbStream = Nothing
        mFileName = ""
        mToWrite = False
        mOpen = False
        CloseFile = True
    End If
End Function


Function GetLine(Optional Trimmed As Boolean = False) As String
    On Error Resume Next
    If mOpen Then
        If mBufferSize Then
            GetLine = mLineBuffer(mBufferSize)
            mBufferSize = mBufferSize - 1
        Else
            Dim s As String
            s = mAdodbStream.ReadText(adReadLine)
            If Trimmed Then
                GetLine = Trim$(s)
            Else
                GetLine = s
            End If
        End If
    Else
        Err.Raise 9999, "clsTextFile.GetLine", "File not open."
    End If
End Function

Sub PutLine(line As String)
' Put a line into the file.
' If the file was opened for reading, this will put the line into a buffer.
' If the file was opened for writing, this will put the line into the stream object.
    If mOpen Then
        If mToWrite Then
            mAdodbStream.WriteText line, adWriteLine
        Else
            mBufferSize = mBufferSize + 1
            If mBufferSize > UBound(mLineBuffer) Then
                ReDim Preserve mLineBuffer(1 To mBufferSize + BUFFER_INCREMENT)
            End If
            mLineBuffer(mBufferSize) = line
        End If
    Else
        Err.Raise 9999, "clsTextFile.PutLine", "File not open."
    End If
End Sub

Private Sub Class_Initialize()
    ReDim mLineBuffer(1 To BUFFER_INCREMENT)
End Sub

Private Sub Class_Terminate()
    If mOpen Then mAdodbStream.Close
End Sub
