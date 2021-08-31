Attribute VB_Name = "mdl_Round"
Option Explicit
Option Private Module

Public Const inoRoundF = 1
Public Const inoRoundU = 2
Public Const inoRoundD = 3


Public strRound(3) As String

Public Sub initRound()
    strRound(1) = "=ROUND("
    strRound(2) = "=ROUNDUP("
    strRound(3) = "=ROUNDDOWN("
End Sub

Public Sub Rounding(ByVal intType As Integer, ByVal intDigits As Integer)
    Dim rng As Range
    Dim c As Range
    Dim blnLarge As Boolean
    Dim blnScreenUpdate As Boolean
    Dim blnDisplayStatusbar As Boolean
    Dim intCalc As Integer
    Dim lngCount As Long
    Dim lngActual As Long
    
    initRound
    Set rng = Selection
    
    blnScreenUpdate = Application.ScreenUpdating
    intCalc = Application.Calculation
    blnDisplayStatusbar = Application.DisplayStatusBar
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    lngCount = rng.Cells.Count
    If lngCount > 50 Then
        MsgBox strError(3), , strError(2)
    End If
    
    lngActual = 1
    For Each c In rng
        Application.StatusBar = strError(4) & lngActual & strError(5) & lngCount
        AddFunction c, intType, intDigits
        lngActual = lngActual + 1
    Next
    
    Application.StatusBar = False
    Application.DisplayStatusBar = blnDisplayStatusbar
    Application.ScreenUpdating = blnScreenUpdate
    Application.Calculation = intCalc
    Application.Calculate

End Sub

Sub AddFunction(ByVal rng As Range, Optional ByVal intType As Integer = 1, Optional ByVal intDigits As Integer = 2)
    Dim strFormula As String
    Dim strFormulaNew As String
    Dim strTest As String
    
    'check if function exist
    strFormula = rng.Formula
    If VBA.Left(strFormula, 1) = "=" Then
        If VBA.InStr(strFormula, "(") > 2 Then
            strTest = VBA.Left(strFormula, VBA.InStr(strFormula, "("))
            Select Case strTest
                Case "=ROUND("
                    strFormulaNew = ReplaceRound(strFormula, inoRoundF, intType, intDigits)
                Case "=ROUNDUP("
                    strFormulaNew = ReplaceRound(strFormula, inoRoundU, intType, intDigits)
                Case "=ROUNDDOWN("
                    strFormulaNew = ReplaceRound(strFormula, inoRoundD, intType, intDigits)
                Case Else
                    'no round
                    strFormulaNew = strRound(intType) & VBA.Mid(strFormula, 2) & ", " & intDigits & ")"
            End Select
        Else
            'no function at start
            strFormulaNew = strRound(intType) & VBA.Mid(strFormula, 2) & ", " & intDigits & ")"
        End If
        rng.Formula = strFormulaNew
    Else
        'no formula
        If IsNumeric(strFormula) Then
            If blnNumbers Then
'            If MsgBox("Soll '" & strFormula & "' gerundet werden?", vbYesNo) = vbYes Then
                rng.Formula = strRound(intType) & strFormula & ", " & intDigits & ")"
            End If
        End If
    End If
    
End Sub

Private Function CountLetters(ByVal strTest As String, ByVal strSearch As String, Optional ByVal blnIgnoreCase As Boolean = True) As Integer
    If blnIgnoreCase Then
        CountLetters = VBA.Len(strTest) - VBA.Len(VBA.Replace(UCase(strTest), UCase(strSearch), ""))
    Else
        CountLetters = VBA.Len(strTest) - VBA.Len(VBA.Replace(strTest, strSearch, ""))
    End If
End Function

Function ReplaceRound(ByVal strFormula As String, ByVal intTypeOld As Integer, ByVal intTypeNew As Integer, ByVal intDigits As Integer) As String
    Dim strFormulaNew As String
    
    If CountLetters(strFormula, "(") = 1 Then
                          
    ElseIf CountLetters(strFormula, "(") = CountLetters(strFormula, ")") Then
    
    Else
    
    End If
    If VBA.InStrRev(strFormula, ",") > 1 Then
        strFormula = VBA.Left(strFormula, VBA.InStrRev(strFormula, ",")) & intDigits & ")"
    Else
        strFormula = VBA.Replace(strFormula, ")", ", " & intDigits & ")")
    End If
    
    strFormulaNew = VBA.Replace(strFormula, strRound(intTypeOld), strRound(intTypeNew))
    ReplaceRound = strFormulaNew
End Function

Public Sub RemoveRounding()
    Dim rng As Range
    Dim c As Range
    Dim blnLarge As Boolean
    Dim blnScreenUpdate As Boolean
    Dim blnDisplayStatusbar As Boolean
    Dim intCalc As Integer
    Dim lngCount As Long
    Dim lngActual As Long
    
    initRound
    Set rng = Selection
    
    blnScreenUpdate = Application.ScreenUpdating
    intCalc = Application.Calculation
    blnDisplayStatusbar = Application.DisplayStatusBar
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    lngCount = rng.Cells.Count
    If lngCount > 50 Then
        MsgBox strError(3), , strError(2)
    End If
    
    lngActual = 1
    For Each c In rng
        Application.StatusBar = strError(4) & lngActual & strError(5) & lngCount
        RemoveFunction c
        lngActual = lngActual + 1
    Next
    
    Application.StatusBar = False
    Application.DisplayStatusBar = blnDisplayStatusbar
    Application.ScreenUpdating = blnScreenUpdate
    Application.Calculation = intCalc
    Application.Calculate

End Sub

Sub RemoveFunction(rng As Range)
    Dim strFormula As String
    Dim strFormulaNew As String
    Dim strTest As String
    'check if function exist
    strFormula = rng.Formula
    If VBA.Left(strFormula, 1) = "=" Then
        strFormulaNew = strFormula
        If VBA.Left(strFormula, 6) = "=ROUND" Then
            strTest = VBA.Left(strFormula, VBA.InStr(strFormula, "("))
            Select Case VBA.InStr(strFormula, "(")
                Case 7 '"=ROUND("
                    strFormulaNew = RemoveRound(strFormula, inoRoundF)
                Case 9 '"=ROUNDUP("
                    strFormulaNew = RemoveRound(strFormula, inoRoundU)
                Case 11 '"=ROUNDDOWN("
                    strFormulaNew = RemoveRound(strFormula, inoRoundD)
                Case Else
                    'no round
            End Select
            
        End If
        rng.Formula = strFormulaNew
    Else
        'no formula
    End If
    
End Sub

Function RemoveRound(ByVal strFormula As String, ByVal intTypeOld As Integer) As String
    Dim strFormulaNew As String
    
    If VBA.InStrRev(strFormula, ",") > 1 Then
        strFormula = VBA.Left(strFormula, VBA.InStrRev(strFormula, ",") - 1)
    Else
        strFormula = VBA.Replace(strFormula, ")", "")
    End If
    
    strFormulaNew = VBA.Replace(strFormula, strRound(intTypeOld), "=")
    
    If IsNumeric(VBA.Mid(strFormulaNew, 2)) Then
        strFormulaNew = VBA.Mid(strFormulaNew, 2)
    End If
    RemoveRound = strFormulaNew
End Function
