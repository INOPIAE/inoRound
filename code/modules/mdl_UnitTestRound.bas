Attribute VB_Name = "mdl_UnitTestRound"
Option Explicit
Option Private Module
' to use this module the com addin rubberduck needs be installed
' https://github.com/rubberduck-vba/Rubberduck

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

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
End Sub

'@TestCleanup
Private Sub TestTerminierung()
    'this method runs after every test in the module.
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
