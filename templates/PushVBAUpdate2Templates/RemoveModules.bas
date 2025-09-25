Attribute VB_Name = "RemoveModules"
Option Explicit
Sub RemoveAllVBA_ModulesFromDestinationWorkbook(destinationWorkbook As Workbook)

'This will remove all modules including ClassModules and UserForms but keep all
'object modules (sheets, workbook).

Dim module As Object

On Error Resume Next
For Each module In destinationWorkbook.VBProject.VBComponents
    destinationWorkbook.VBProject.VBComponents.Remove module
Next
On Error GoTo 0

End Sub
