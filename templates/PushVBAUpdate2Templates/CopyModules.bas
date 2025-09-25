Attribute VB_Name = "CopyModules"
Option Explicit
Sub CopyAllVBA_ModulesFromSourceWorkbookToDestinationWorkbook(sourceWorkbook As Workbook, destinationWorkbook As Workbook)

'This will copy from source workbook all modules except sheets, workbook.
'As direct copy is not allowed it is accomplished by 2 steps
'1 module from sourceWorkbook is exported into temporary file
'2 temporary file is imported into destinationWorkbook and then temporary file is deleted

Dim module As Object, pathToTemporaryFilesFolder As String, pathToTemporaryFile As String

pathToTemporaryFilesFolder = "C:\Users\Jerrison Chai\Documents\02 DEMO\TempVBAmacro"

On Error Resume Next
For Each module In sourceWorkbook.VBProject.VBComponents
    If InStr(module.Name, "ThisWorkbook") = 0 And InStr(module.Name, "Sheet") = 0 Then
        pathToTemporaryFile = pathToTemporaryFilesFolder & "\" & module.Name & ".bas"
        module.Export (pathToTemporaryFile)
        destinationWorkbook.VBProject.VBComponents.Import (pathToTemporaryFile)
        Kill pathToTemporaryFile
    End If
Next
On Error GoTo 0

End Sub
