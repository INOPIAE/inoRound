Attribute VB_Name = "MWindows"
'
' Description:  Contains API constants, variables, declarations and procedures
'               to demonstrate API routines related to the windows
'
' Authors:      Stephen Bullen, www.oaltd.co.uk
'               Rob Bovey, www.appspro.com
'
' ==
' == 24-Jan-11 preliminarily 64bit safe by Daniel Kraus
' == 05-Nov-11 GetText support


Option Explicit
Option Private Module

' **************************************************************
' Declarations for the ApphWnd and FindOurWindow example functions
' **************************************************************

#If VBA7 Then
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
        (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
#Else
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
#End If



' **************************************************************
' Declarations for the WorkbookWindowhWnd example function
' The WorkbookWindowhWnd procedure uses FindWindowEx, defined above
' **************************************************************


' **************************************************************
' Declarations for the SetNameDropdownWidth example procedure
' **************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''
' Constants used in the SendMessage call
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const CB_SETDROPPEDWIDTH As Long = &H160&     'from winuser.h

''''''''''''''''''''''''''''''''''''''''''''''''''
' Function Declarations
' The SetNameDropdownWidth procedure also uses FindWindowEx, defined above
''''''''''''''''''''''''''''''''''''''''''''''''''
'Send a message to a window
#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If


' **************************************************************
' Declarations for the SetIcon example procedure
' **************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''
' Constants used in the SendMessage call
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const WM_SETICON As Long = &H80

''''''''''''''''''''''''''''''''''''''''''''''''''
' Function Declarations
' The SetIcon procedure also uses SendMessage, defined above
''''''''''''''''''''''''''''''''''''''''''''''''''
'Get a handle to an icon from a file
#If VBA7 Then
    Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" _
        (ByVal hInst As LongPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As LongPtr
#Else
    Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
#End If




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Foolproof way to find the main Excel window handle
'
' Arguments:    None
'
' Returns:      Long        The handle of Excel's main window
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'
#If VBA7 Then
    Function ApphWnd() As LongPtr
#Else
    Function ApphWnd() As Long
#End If

    If Application.Version >= 10 Then
        ApphWnd = Application.hwnd
    Else
        ApphWnd = FindOurWindow("XLMAIN", Application.Caption)
    End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Finds a top-level window of the given class and
'           caption that belongs to this instance of Excel,
'           by matching the process IDs
'
' Arguments:    sClass      The window class name to look for
'               sCaption    The window caption to look for
'
' Returns:      Long        The handle of Excel's main window
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'

#If VBA7 Then
    Function FindOurWindow(Optional ByVal sClass As String = vbNullString, _
        Optional ByVal sCaption As String = vbNullString) As LongPtr
    Dim hWndDesktop As LongPtr
    Dim hwnd As LongPtr
    Dim hProcThis As LongPtr
#Else
    Function FindOurWindow(Optional ByVal sClass As String = vbNullString, _
        Optional ByVal sCaption As String = vbNullString) As Long
    Dim hWndDesktop As Long
    Dim hwnd As Long
    Dim hProcThis As Long
#End If
    Dim hProcWindow As Long

    'All top-level windows are children of the desktop,
    'so get that handle first
    hWndDesktop = GetDesktopWindow

    'Get the ID of this instance of Excel, to match
    hProcThis = GetCurrentProcessId

    Do
        'Find the next child window of the desktop that
        'matches the given window class and/or caption.
        'The first time in, hWnd will be zero, so we'll get
        'the first matching window.  Each call will pass the
        'handle of the window we found the last time, thereby
        'getting the next one (if any)
        hwnd = FindWindowEx(hWndDesktop, hwnd, sClass, sCaption)

        'Get the ID of the process that owns the window we found
        GetWindowThreadProcessId hwnd, hProcWindow

        'Loop until the window's process matches this process,
        'or we didn't find the window
    Loop Until hProcWindow = hProcThis Or hwnd = 0

    'Return the handle we found
    FindOurWindow = hwnd

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Find the handle of a given workbook's Window
'
' Arguments:    None
'
' Returns:      Long        The handle of the given Window
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'
#If VBA7 Then
    Function WorkbookWindowhWnd(ByRef wndWindow As Window) As LongPtr
    
        Dim hWndExcel As LongPtr
        Dim hWndDesk As LongPtr
    
        'Get the main Excel window
        hWndExcel = ApphWnd
    
        'Find the desktop
        hWndDesk = FindWindowEx(hWndExcel, 0, "XLDESK", vbNullString)
    
        'Find the workbook window
        WorkbookWindowhWnd = FindWindowEx(hWndDesk, 0, "EXCEL7", wndWindow.Caption)
    
    End Function
#Else
    Function WorkbookWindowhWnd(ByRef wndWindow As Window) As Long
    
        Dim hWndExcel As Long
        Dim hWndDesk As Long
    
        'Get the main Excel window
        hWndExcel = ApphWnd
    
        'Find the desktop
        hWndDesk = FindWindowEx(hWndExcel, 0, "XLDESK", vbNullString)
    
        'Find the workbook window
        WorkbookWindowhWnd = FindWindowEx(hWndDesk, 0, "EXCEL7", wndWindow.Caption)
    
    End Function
#End If



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Make the Name dropdown list 200 pixels wide
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'
Sub SetNameDropdownWidth()

    #If VBA7 Then
        Dim hWndExcel As LongPtr
        Dim hWndFormulaBar As LongPtr
        Dim hWndNameCombo As LongPtr
    #Else
        Dim hWndExcel As Long
        Dim hWndFormulaBar As Long
        Dim hWndNameCombo As Long
    #End If


    'Get the main Excel window
    hWndExcel = ApphWnd

    'Get the handle for the formula bar window
    hWndFormulaBar = FindWindowEx(hWndExcel, 0, "EXCEL;", vbNullString)

    'Get the handle for the Name combobox
    hWndNameCombo = FindWindowEx(hWndFormulaBar, 0, "combobox", vbNullString)

    'Set the dropdown list to be 200 pixels wide
    SendMessage hWndNameCombo, CB_SETDROPPEDWIDTH, 200, 0

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Set a window's icon
'           When workbook windows are maximised, Excel doesn't update
'           the icon (shown on the left of the menu bar) until the window
'           is minimised/restored.
'
' Arguments:    hwnd        The handle of the window to change the icon for
'               sIcon       The path of the icon file
'
' Date          Developer       Action
' --------------------------------------------------------------
' 02 Jun 04     Stephen Bullen  Created
'
#If VBA7 Then
    Sub SetIcon(ByVal hwnd As LongPtr, ByVal sIcon As String)
    
        Dim hIcon As LongPtr
    
        'Get the icon handle
        hIcon = ExtractIcon(0, sIcon, 0)
    
        'Set the big (32x32) and small (16x16) icons
        SendMessage hwnd, WM_SETICON, True, hIcon
        SendMessage hwnd, WM_SETICON, False, hIcon
    
    End Sub
#Else
    Sub SetIcon(ByVal hwnd As Long, ByVal sIcon As String)
    
        Dim hIcon As Long
    
        'Get the icon handle
        hIcon = ExtractIcon(0, sIcon, 0)
    
        'Set the big (32x32) and small (16x16) icons
        SendMessage hwnd, WM_SETICON, True, hIcon
        SendMessage hwnd, WM_SETICON, False, hIcon
    
    End Sub
#End If
