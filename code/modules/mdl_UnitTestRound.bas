Attribute VB_Name = "mdl_UnitTestRound"
Option Explicit
Option Private Module
' to use this module the COM add-in rubberduck needs be installed
' https://github.com/rubberduck-vba/Rubberduck

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private wks As Worksheet

'@ModuleInitialize
Private Sub ModulInitialisierung()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    initRound
End Sub

'@ModuleCleanup
Private Sub ModulTerminierung()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialisierung()
    'This method runs before every test in the module..
    Set wks = ActiveWorkbook.Worksheets.Add
End Sub

'@TestCleanup
Private Sub TestTerminierung()
    'this method runs after every test in the module.
    Dim blnAlert As Boolean
    blnAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wks.Delete
    Application.DisplayAlerts = blnAlert
End Sub


'@TestMethod("Uncategorized")
Private Sub TestReplaceRound()
    On Error GoTo TestFail:
    
    ' function to be tested:
    ' ReplaceRound(ByVal strFormula As String, ByVal intTypeOld As Integer, ByVal intTypeNew As Integer, ByVal intDigits As Integer) As String
    
    'Arrange:
    Dim strFormula As String
    Dim intTypeOld As Integer
    Dim intTypeNew As Integer
    Dim intDigits As Integer
    Dim strResult As String
    Dim strExpected As String
    Dim intCount As Integer

    'Act:
        
    'Assert:
    strFormula = "=ROUND(A2,1)"
    For intCount = 1 To 3
        strResult = ReplaceRound(strFormula, inoRoundF, intCount, 3)
        strExpected = strRound(intCount) & "A2,3)"
        Assert.AreEqual strExpected, strResult
    Next
    
    strFormula = "=ROUNDDOWN(A2,1)"
    For intCount = 1 To 3
        strResult = ReplaceRound(strFormula, inoRoundD, intCount, 3)
        strExpected = strRound(intCount) & "A2,3)"
        Assert.AreEqual strExpected, strResult
    Next
    
    strFormula = "=ROUNDUP(A2,1)"
    For intCount = 1 To 3
        strResult = ReplaceRound(strFormula, inoRoundU, intCount, 3)
        strExpected = strRound(intCount) & "A2,3)"
        Assert.AreEqual strExpected, strResult
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestRemoveRound()
    On Error GoTo TestFail
    
    ' function to be tested:
    ' RemoveRound(ByVal strFormula As String, ByVal intTypeOld As Integer) As String
    
    'Arrange:
    Dim strFormula As String
    Dim intTypeOld As Integer
    Dim strResult As String
    Dim strExpected As String
    Dim intCount As Integer
    
    'Act:
    
    'Assert:
    'check reference
    For intCount = 1 To 3
        strFormula = strRound(intCount) & "A2,1)"
        strResult = RemoveRound(strFormula, intCount)
        strExpected = "=A2"
        Assert.AreEqual strExpected, strResult
    Next
    
    'check formula
    For intCount = 1 To 3
        strFormula = strRound(intCount) & "SUM(A2:A3),1)"
        strResult = RemoveRound(strFormula, intCount)
        strExpected = "=SUM(A2:A3)"
        Assert.AreEqual strExpected, strResult
    Next
    
    'check figure
    For intCount = 1 To 3
        strFormula = strRound(intCount) & "47.11,1)"
        strResult = RemoveRound(strFormula, intCount)
        strExpected = "47.11"
        Assert.AreEqual strExpected, strResult
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub FillWorksheet()
    With wks
        .Range("A1").Formula = "6"
        .Range("A2").Formula = "=A1*11%"
        .Range("A3").Formula = "=ROUND(A2,1)"
        .Range("A4").Formula = "=ROUNDUP(A2,1)"
        .Range("A5").Formula = "=ROUNDDOWN(A2,1)"
        .Range("A6").Formula = "=SUM(A1:A2)"
        .Range("A7").Formula = "AB"
        .Range("A8").Formula = "=ROUND(10.345,1)"
        .Range("A9").Formula = "=ROUND(Sum(A1:A2),1)"
    End With

End Sub

'@TestMethod("Uncategorized")
Private Sub TestReplaceRounding()
    On Error GoTo TestFail:
    
    ' sub to be tested:
    ' Rounding(ByVal intType As Integer, ByVal intDigits As Integer)
    
    'Arrange:
    Dim intCount As Integer
    Dim intDecimal As Integer
    Dim intTest As Integer
    Dim strResult As String
    Dim strExpected As String
    

    intDecimal = 2
    
    'Act:
        
    'Assert:
    
    For intTest = 1 To 3
        FillWorksheet
        wks.Range("A1:A9").Select
        Rounding intTest, intDecimal
    
        For intCount = 1 To 9
            Select Case intCount
                Case 1
                    strExpected = "6"
                Case 2
                    strExpected = strRound(intTest) & "A1*11%, " & intDecimal & ")"
                Case 3
                    strExpected = strRound(intTest) & "A2," & intDecimal & ")"
                Case 4
                    strExpected = strRound(intTest) & "A2," & intDecimal & ")"
                Case 5
                    strExpected = strRound(intTest) & "A2," & intDecimal & ")"
                Case 6
                    strExpected = strRound(intTest) & "SUM(A1:A2), " & intDecimal & ")"
                Case 7
                    strExpected = "AB"
                Case 8
                    strExpected = strRound(intTest) & "10.345," & intDecimal & ")"
                Case 9
                    strExpected = strRound(intTest) & "SUM(A1:A2)," & intDecimal & ")"
            End Select
            strResult = wks.Cells(intCount, 1).Formula
            Assert.AreEqual strExpected, strResult
        Next
    Next
    
    For intTest = 1 To 3
        FillWorksheet
        wks.Range("A2,A4").Select
        Rounding intTest, intDecimal
    
        For intCount = 1 To 9
            Select Case intCount
                Case 1
                    strExpected = "6"
                Case 2
                    strExpected = strRound(intTest) & "A1*11%, " & intDecimal & ")"
                Case 3
                    strExpected = "=ROUND(A2,1)"
                Case 4
                    strExpected = strRound(intTest) & "A2," & intDecimal & ")"
                Case 5
                    strExpected = "=ROUNDDOWN(A2,1)"
                Case 6
                    strExpected = "=SUM(A1:A2)"
                Case 7
                    strExpected = "AB"
                Case 8
                    strExpected = "=ROUND(10.345,1)"
                Case 9
                    strExpected = "=ROUND(SUM(A1:A2),1)"
            End Select
            strResult = wks.Cells(intCount, 1).Formula
            Assert.AreEqual strExpected, strResult
        Next
    Next

    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestRemoveRounding()
    On Error GoTo TestFail:
    
    ' sub to be tested:
    ' RemoveRounding
    
    'Arrange:
    Dim intCount As Integer
    Dim strResult As String
    Dim strExpected As String
    
    'Act:
        
    'Assert:
    
    FillWorksheet
    wks.Range("A1:A9").Select
    RemoveRounding

    For intCount = 1 To 9
        Select Case intCount
            Case 1
                strExpected = "6"
            Case 2
                strExpected = "=A1*11%"
            Case 3
                strExpected = "=A2"
            Case 4
                strExpected = "=A2"
            Case 5
                strExpected = "=A2"
            Case 6
                strExpected = "=SUM(A1:A2)"
            Case 7
                strExpected = "AB"
            Case 8
                strExpected = "10.345"
            Case 9
                strExpected = "=SUM(A1:A2)"
        End Select
        strResult = wks.Cells(intCount, 1).Formula
        Assert.AreEqual strExpected, strResult
    Next


    FillWorksheet
    wks.Range("A2,A4").Select
    RemoveRounding

    For intCount = 1 To 9
        Select Case intCount
            Case 1
                strExpected = "6"
            Case 2
                strExpected = "=A1*11%"
            Case 3
                strExpected = "=ROUND(A2,1)"
            Case 4
                strExpected = "=A2"
            Case 5
                strExpected = "=ROUNDDOWN(A2,1)"
            Case 6
                strExpected = "=SUM(A1:A2)"
            Case 7
                strExpected = "AB"
            Case 8
                strExpected = "=ROUND(10.345,1)"
            Case 9
                strExpected = "=ROUND(SUM(A1:A2),1)"
        End Select
        strResult = wks.Cells(intCount, 1).Formula
        Assert.AreEqual strExpected, strResult
    Next
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestReplaceRoundingNumber()
    On Error GoTo TestFail:
    
    ' sub to be tested:
    ' Rounding(ByVal intType As Integer, ByVal intDigits As Integer)
    
    'Arrange:
    Dim intCount As Integer
    Dim intDecimal As Integer
    Dim intTest As Integer
    Dim strResult As String
    Dim strExpected As String
    Dim blnNumber As Boolean
    
    blnNumber = blnNumbers
    blnNumbers = True
    intDecimal = 2
    
    'Act:
        
    'Assert:
    
    For intTest = 1 To 3
        FillWorksheet
        wks.Range("A1:A9").Select
        Rounding intTest, intDecimal
    
        For intCount = 1 To 9
            Select Case intCount
                Case 1
                    strExpected = strRound(intTest) & "6, " & intDecimal & ")"
                Case 2
                    strExpected = strRound(intTest) & "A1*11%, " & intDecimal & ")"
                Case 3
                    strExpected = strRound(intTest) & "A2," & intDecimal & ")"
                Case 4
                    strExpected = strRound(intTest) & "A2," & intDecimal & ")"
                Case 5
                    strExpected = strRound(intTest) & "A2," & intDecimal & ")"
                Case 6
                    strExpected = strRound(intTest) & "SUM(A1:A2), " & intDecimal & ")"
                Case 7
                    strExpected = "AB"
                Case 8
                    strExpected = strRound(intTest) & "10.345," & intDecimal & ")"
                Case 9
                    strExpected = strRound(intTest) & "SUM(A1:A2)," & intDecimal & ")"
            End Select
            strResult = wks.Cells(intCount, 1).Formula
            Assert.AreEqual strExpected, strResult
        Next
    Next
    
    blnNumbers = blnNumber
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
