Attribute VB_Name = "mdl_Round"
Option Explicit

Public Const inoRoundF = 1
Public Const inoRoundU = 2
Public Const inoRoundD = 3


Private strRound(3) As String

Private Sub initRound()
    strRound(1) = "=ROUND("
    strRound(2) = "=ROUNDUP("
    strRound(3) = "=ROUNDDOWN("
End Sub

Public Sub Rounding(ByVal intType As Integer, ByVal intDigits As Integer)
    Dim rng As Range
    Dim c As Range
    initRound
    Set rng = Selection
    For Each c In rng
        AddFunction c, intType, intDigits
    Next
End Sub

Sub AddFunction(rng As Range, Optional ByVal intType As Integer = 1, Optional ByVal intDigits As Integer = 2)
    Dim strFormula As String
    Dim strFormulaNew As String
    Dim strTest As String
    'check if function exist
    strFormula = rng.Formula
    If Left(strFormula, 1) = "=" Then
        If InStr(strFormula, "(") > 2 Then
            strTest = Left(strFormula, InStr(strFormula, "("))
            Select Case strTest
                Case "=ROUND("
                    strFormulaNew = ReplaceRound(strFormula, inoRoundF, intType, intDigits)
                Case "=ROUNDUP("
                    strFormulaNew = ReplaceRound(strFormula, inoRoundU, intType, intDigits)
                Case "=ROUNDDOWN("
                    strFormulaNew = ReplaceRound(strFormula, inoRoundD, intType, intDigits)
                Case Else
                    'no round
                    strFormulaNew = strRound(intType) & Mid(strFormula, 2) & ", " & intDigits & ")"
            End Select
        Else
            'no function at start
            strFormulaNew = strRound(intType) & Mid(strFormula, 2) & ", " & intDigits & ")"
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
        CountLetters = Len(strTest) - Len(Replace(UCase(strTest), UCase(strSearch), ""))
    Else
        CountLetters = Len(strTest) - Len(Replace(strTest, strSearch, ""))
    End If
End Function

Function ReplaceRound(ByVal strFormula As String, ByVal intTypeOld As Integer, ByVal intTypeNew As Integer, ByVal intDigits As Integer) As String
    Dim strFormulaNew As String
    
    If CountLetters(strFormula, "(") = 1 Then
                          
    ElseIf CountLetters(strFormula, "(") = CountLetters(strFormula, ")") Then
    
    Else
    
    End If
    If InStrRev(strFormula, ",") > 1 Then
        strFormula = Left(strFormula, InStrRev(strFormula, ",")) & intDigits & ")"
    Else
        strFormula = Replace(strFormula, ")", ", " & intDigits & ")")
    End If
    
    strFormulaNew = Replace(strFormula, strRound(intTypeOld), strRound(intTypeNew))
    ReplaceRound = strFormulaNew
End Function

Sub test()
    With ActiveSheet
        .Range("A1").Formula = "6"
        .Range("A2").Formula = "=A1*11%"
        .Range("A3").Formula = "=ROUND(A2,1)"
        .Range("A4").Formula = "=ROUNDUP(A2,1)"
        .Range("A5").Formula = "=ROUNDDOWN(A2,1)"
        .Range("A6").Formula = "=Sum(A1:A2)"
        .Range("A7").Formula = "AB"
       
        .Range("A1:A7").Select
    End With

    'Rounding
End Sub


