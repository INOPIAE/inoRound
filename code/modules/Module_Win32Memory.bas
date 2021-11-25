Attribute VB_Name = "Module_Win32Memory"
' ==================================================================
' == XL TOOLBOX   (c) 2008-2013 Daniel Kraus   Licensed under GPLv2
' ==================================================================
' == DanielsXLToolbox.Module_Win32Memory
' ==
' == Created: 09-Feb-10 12:45
' ==
' ==
' ==
' == 24-Jan-11 preliminarily 64bit safe
' == 05-Nov-11 GetText support
' ==



Option Explicit
Option Private Module

    Type MEMORYSTATUSEX
       dwLength As Long
       dwMemoryLoad As Long
       ullTotalPhys As DWORDLONG
       ullAvailPhys As DWORDLONG
       ullTotalPageFile As DWORDLONG
       ullAvailPageFile As DWORDLONG
       ullTotalVirtual As DWORDLONG
       ullAvailVirtual As DWORDLONG
       ullAvailExtendedVirtual As DWORDLONG
    End Type


#If VBA7 Then
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
    Declare PtrSafe Sub CopyMemoryWrite Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal Destination As LongPtr, Source As Any, ByVal Length As LongPtr)
    Declare PtrSafe Sub CopyMemoryRead Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)
    Declare PtrSafe Function lstrlenW Lib "kernel32" _
        (ByVal lpString As Long) As Long
    Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenA" _
        (ByVal lpString As Any) As Long
    Declare PtrSafe Function GlobalLock Lib "kernel32" _
        (ByVal hMEM As LongPtr) As LongPtr
    Declare PtrSafe Function GetProcessHeap Lib "kernel32" () As LongPtr
    Declare PtrSafe Function HeapAlloc Lib "kernel32" _
        (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Declare PtrSafe Function HeapFree Lib "kernel32" _
        (ByVal hHeap As LongPtr, ByVal dwFlags As Long, lpMem As Any) As Long
    Declare PtrSafe Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" _
        (ByVal hFile As LongPtr, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, _
        ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As LongPtr
    Private Declare PtrSafe Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    Declare Sub CopyMemoryWrite Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Any, ByVal Length As Long)
    Declare Sub CopyMemoryRead Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
    Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
    Declare Function GlobalLock Lib "kernel32" (ByVal hMEM As Long) As Long
    Declare Function GetProcessHeap Lib "kernel32" () As Long
    Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
    Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
    Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, _
        ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
    Private Declare Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX)
#End If





Function GetMemStatus() As MEMORYSTATUSEX
    Dim ms As MEMORYSTATUSEX

    'Set the length member before you call GlobalMemoryStatus
    ms.dwLength = Len(ms)
    GlobalMemoryStatusEx ms
    GetMemStatus = ms
    
End Function


Function GetTotalMemory(ms As MEMORYSTATUSEX, Optional Dimension As Long = 0) As Double
' Dimension can be supplied as follows:
'   0 - return bytes
'   1 - return kilobytes (1024^1 bytes)
'   2 - return megabytes (1024^2 bytes)
'   3 - return gigabytes (1024^3 bytes)

    ' Turn off runtime error handling as this function may be called
    ' by the error handler itself.
    On Error Resume Next
    With ms
        GetTotalMemory = DWordLongToDouble(.ullTotalPhys) / (1024 ^ Dimension)
    End With
End Function


Function GetTotalVirtualMemory(ms As MEMORYSTATUSEX, Optional Dimension As Long = 0) As Double
' Dimension can be supplied as follows:
'   0 - return bytes
'   1 - return kilobytes (1024^1 bytes)
'   2 - return megabytes (1024^2 bytes)
'   3 - return gigabytes (1024^3 bytes)

    ' Turn off runtime error handling as this function may be called
    ' by the error handler itself.
    On Error Resume Next
    With ms
        GetTotalVirtualMemory = DWordLongToDouble(.ullTotalVirtual) / 1024 ^ Dimension
    End With
End Function


Function GetAvailableMemory(ms As MEMORYSTATUSEX, Optional Dimension As Long = 0) As Double
' Dimension can be supplied as follows:
'   0 - return bytes
'   1 - return kilobytes (1024^1 bytes)
'   2 - return megabytes (1024^2 bytes)
'   3 - return gigabytes (1024^3 bytes)
    
    ' Turn off runtime error handling as this function may be called
    ' by the error handler itself.
    On Error Resume Next
    With ms
        GetAvailableMemory = DWordLongToDouble(.ullAvailPhys) / 1024 ^ Dimension
    End With
End Function


Function GetAvailableVirtualMemory(ms As MEMORYSTATUSEX, Optional Dimension As Long = 0) As Double
' Dimension can be supplied as follows:
'   0 - return bytes
'   1 - return kilobytes (1024^1 bytes)
'   2 - return megabytes (1024^2 bytes)
'   3 - return gigabytes (1024^3 bytes)
    
    ' Turn off runtime error handling as this function may be called
    ' by the error handler itself.
    On Error Resume Next
    
    With ms
        GetAvailableVirtualMemory = DWordLongToDouble(.ullAvailVirtual) / 1024 ^ Dimension
    End With
End Function


Function DWordLongToDouble(dwl As DWORDLONG) As Double
' Because VBA cannot handle 8-byte integers (leave alone
' 4-byte unsigned integers) we need to use a trick to
' calculate a true 64-bit integer value from two 4-byte
' integers: Bit shift right by 1 bit, multiply the result
' as a double and add back the low bit; multiply by 2^32
' and add the corresponding value for the lower 32 bit
' of the integer. It's slow, but it works.
    On Error Resume Next
    With dwl
        DWordLongToDouble = _
            2# * BitShiftRight(.Hi, 1) + (.Hi And &H1) * 2 ^ 32 + _
            2# * BitShiftRight(.Lo, 1) + (.Lo And &H1)
    End With
End Function


Function BitShiftRight(l As Long, Shift As Long) As Long
' Performs a bit shift to the right
' Because VBA does not know unsigned long integers,
' we need to employ a trick.

    If Shift = 0 Then
        BitShiftRight = l
        Exit Function
    End If

    If l >= 0 Then
        BitShiftRight = l \ (2 ^ Shift)
    Else
        ' The high bit is set, and VBA interprets
        ' the value as a negative one.
        ' Therefore we ignore the high bit during the
        ' shift and add it back
        ' Mask for everything except the high bit:
        '   &H7FFFFFFF - 0111 1111 1111 1111 1111 1111 1111 1111
        ' Mask for the high bit:
        '   &H80000000 - 1000 0000 0000 0000 0000 0000 0000 0000
        ' Mask for the "second high" bit:
        '   &H40000000 - 0100 0000 0000 0000 0000 0000 0000 0000
        
        BitShiftRight = ((l And &H7FFFFFFF) \ 2 ^ Shift) Or (&H40000000 \ 2 ^ (Shift - 1))
    End If
End Function

