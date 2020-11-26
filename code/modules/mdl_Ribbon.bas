Attribute VB_Name = "mdl_Ribbon"
Option Explicit

Private mDDcurrentID As Variant
Private intCurrentDigits As Integer
Public gobjRibbon As IRibbonUI
Public blnNumbers As Boolean

'Ribbon functions
Public Sub OnRibbonLoad(ByRef ribbon As IRibbonUI)
    Set gobjRibbon = ribbon
End Sub

Public Sub rbGetLabel(ByRef control As IRibbonControl, ByRef label As Variant)
    SetLanguage
    Select Case control.ID
        Case "grpInoRound"
            label = strLabel(0)
        Case "btnRound"
            'label = strLabel(1)
        Case "btnRoundUp"
            'label = strLabel(2)
        Case "btnRoundDown"
            'label = strLabel(3)
        Case "chkNumbers"
            label = strLabel(4)
        Case "mnuInoRound"
            'label = strLabel(5)
        Case "btnInfoInoRound"
            label = strLabel(6)
        Case Else
            label = ""
    End Select
End Sub

Public Sub rbGetScreentip(ByRef control As IRibbonControl, ByRef text)
    Select Case control.ID
        Case "btnRound"
            text = strScreentip(0)
        Case "btnRoundUp"
            text = strScreentip(1)
        Case "btnRoundDown"
            text = strScreentip(2)
        Case "chkNumbers"
            text = strScreentip(3)
        Case "cboDigits"
            text = strScreentip(4)
        Case Else
            text = ""
    End Select
End Sub

Public Sub rbGetSupertip(ByRef control As IRibbonControl, ByRef text)
    Select Case control.ID
        Case "btnRound"
            text = strSupertip(0)
        Case "btnRoundUp"
            text = strSupertip(1)
        Case "btnRoundDown"
            text = strSupertip(2)
        Case "chkNumbers"
            text = strSupertip(3)
        Case "cboDigits"
            text = strSupertip(4)
        Case Else
            text = ""
    End Select
End Sub

' control functions
Public Sub rbRound(ByRef ctrl As IRibbonControl)
    Rounding inoRoundF, intCurrentDigits
End Sub

Public Sub rbRoundUp(ByRef ctrl As IRibbonControl)
    Rounding inoRoundU, intCurrentDigits
End Sub

Public Sub rbRoundDown(ByRef ctrl As IRibbonControl)
    Rounding inoRoundD, intCurrentDigits
End Sub

Public Sub rbDD(ByRef ctrl As IRibbonControl, ByRef dropdownID As String, ByRef selectedIndex As Variant)
    intCurrentDigits = CInt(Mid(dropdownID, 3))
End Sub

Sub rbDD_GetSelectedItemIndex(ByRef ctrl As IRibbonControl, ByRef returnedVal As Variant)
    If IsEmpty(mDDcurrentID) Then
        mDDcurrentID = 0
        intCurrentDigits = 2
    End If
    returnedVal = mDDcurrentID
End Sub

Sub rbChkNumbers(ByRef control As IRibbonControl, _
   ByRef Pressed As Boolean)
    blnNumbers = Pressed
End Sub

Public Sub rbCboDigits(ByRef control As Office.IRibbonControl, _
   ByRef text As Variant)
    Dim strText As String
    strText = Replace(text, Application.International(xlThousandsSeparator), "")
    If IsNumeric(text) Then
        Dim strTest() As String
        strTest = Split(text, Application.International(xlDecimalSeparator))
        If UBound(strTest) = 0 Then
            intCurrentDigits = -1 * (Len(strTest(0)) - 1)
        Else
            intCurrentDigits = Len(strTest(1))
        End If
    Else
        MsgBox strError(1), , strError(0)
    End If
End Sub

Public Sub rbCboDigits_GetItemLabel( _
   ByRef control As Office.IRibbonControl, _
   ByRef index As Integer, _
   ByRef ItemLabel As Variant)
   
   Select Case index
        Case 0
            ItemLabel = "0" & Application.International(xlDecimalSeparator) & "01"
        Case 1
            ItemLabel = "0" & Application.International(xlDecimalSeparator) & "1"
        Case 2
            ItemLabel = "1"
        Case 3
            ItemLabel = "10"
        Case 4
            ItemLabel = "100"
        Case 5
            ItemLabel = "1000"
        Case Else
            ItemLabel = ""
    End Select
End Sub

Public Sub rbCboDigits_Count(ByRef control As Office.IRibbonControl, _
   ByRef Count As Variant)
   Count = 6
End Sub

Public Sub rbCboDigits_GetText(ByRefcontrol As IRibbonControl, ByRef text)
    text = "0" & Application.International(xlDecimalSeparator) & "01"
End Sub

Public Sub rbInfoInoRound(ByRef control As IRibbonControl)
    frm_Info.Show
End Sub

Private Sub SetLanguage()
    Dim lc As Long
    lc = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    Select Case lc
        Case 1031
            germanText
        Case 1033
            englishText
        Case Else
            englishText
    End Select
    
End Sub

Sub rbTest(ByRefctrl As IRibbonControl)
    test
End Sub
