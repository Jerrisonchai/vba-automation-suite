Attribute VB_Name = "UpdateCode"
Option Explicit

Sub ControlSelectedWorkbookModulesUpdatingFromSourceWorkbook()
Call capturetime
Call MyShape_Click
Dim startTime As Date
Dim endTime As Date
Dim timetaken As Date
Dim wbkO As Workbook

Set wbkO = ThisWorkbook

startTime = Now()
Dim sourceWorkbook As Workbook, destinationWorkbook As Workbook
Dim pathToSourceWorkbook As String
Dim fileDialog As Office.fileDialog, pathToDestinationWorkbook As String

Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)



With fileDialog
  .AllowMultiSelect = False
  .Title = "Please select the file."
  .Filters.Clear
  .Filters.Add "Excel", "*.xlsm"
  If .Show = True Then
    pathToDestinationWorkbook = fileDialog.SelectedItems(1)
  End If
End With
OptimizedMode True
'Application.ScreenUpdating = False

Set destinationWorkbook = Workbooks.Open(pathToDestinationWorkbook)

Call RemoveModules.RemoveAllVBA_ModulesFromDestinationWorkbook(destinationWorkbook)

pathToSourceWorkbook = wbkO.Sheets("Dashboard").Range("C21").Value

Set sourceWorkbook = Workbooks.Open(pathToSourceWorkbook)

Call CopyModules.CopyAllVBA_ModulesFromSourceWorkbookToDestinationWorkbook(sourceWorkbook, destinationWorkbook)

sourceWorkbook.Close Savechanges:=False
destinationWorkbook.Close Savechanges:=True
Set wbkO = Nothing

OptimizedMode False
'Application.ScreenUpdating = True
endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Success"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
[UserName].Value = Environ("UserName")

Call captureendtime
MsgBox "The end of program", vbInformation, ThisWorkbook.Name

End Sub


