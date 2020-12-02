Attribute VB_Name = "mdl_Version"
Option Explicit

Public Const strMakroVersion As String = "1.02"
Public Const dtVersionOf As Date = #2/12/2020#
Private Const strVBProjects As String = "inoRound"
Private Const strVBComponents As String = "DieseArbeitsmappe"

'Set a reference to "Microsoft Visual Basic For Applications Extensibility 5.3" and to "Microsoft Scripting Runtime"

Public Sub Build()
    UpdateXlamFileProperties
    ExportModules
End Sub

Public Sub UpdateXlamFileProperties()
    'use manualy to update the file properties after code changes
    Dim wkb As Workbook
    Dim strText As String
    Dim intPos As Integer
    
    Application.VBE.VBProjects(strVBProjects).VBComponents(strVBComponents).Properties("IsAddin") = False
    Set wkb = Application.Workbooks(strVBProjects & ".xlam")
'    strText = wkb.BuiltinDocumentProperties("Comments")
    strText = "Version " & strMakroVersion & " " & Format(dtVersionOf, "D. MMMM YYYY")
    wkb.BuiltinDocumentProperties("Comments") = strText
    Application.VBE.VBProjects(strVBProjects).VBComponents(strVBComponents).Properties("IsAddin") = True

End Sub

Private Sub ExportModules()
    ' function from Ron de Bruin taken from https://www.rondebruin.nl/win/s9/win002.htm
    ' and  modified
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
'    On Error Resume Next
'        Kill FolderWithVBAProjectFiles & "\*.*"
'    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ThisWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = "classes\" & szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = "forms\" & szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = "modules\" & szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

End Sub


Public Sub ImportModules()
    ' function from Ron de Bruin taken from https://www.rondebruin.nl/win/s9/win002.htm
    ' and  modified
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

Public Function FolderWithVBAProjectFiles() As String
    ' function from Ron de Bruin taken from https://www.rondebruin.nl/win/s9/win002.htm
    ' and  modified
    
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = ThisWorkbook.Path

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    SpecialPath = SpecialPath & "code\"
    
    Dim strFolders()
    Dim strF
    strFolders = Array("forms", "modules", "classes")
    For Each strF In strFolders
        If FSO.FolderExists(SpecialPath & strF) = False Then
            On Error Resume Next
            MkDir SpecialPath & strF
            On Error GoTo 0
        End If
    Next

    
    If FSO.FolderExists(SpecialPath) = True Then
        FolderWithVBAProjectFiles = SpecialPath
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Private Function DeleteVBAModulesAndUserForms()
    ' function from Ron de Bruin taken from https://www.rondebruin.nl/win/s9/win002.htm
    ' and  modified
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    Set VBProj = ActiveWorkbook.VBProject
    
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Function

