Attribute VB_Name = "ExportData"
Option Explicit
Sub SaveAs()    'Button - Save As
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
ThisWorkbook.Save
OptimizedMode True
Application.DisplayAlerts = False
Dim sFolder As String, FolderPicker As FileDialog, sFile As String, mypath As String, wbkS As Workbook, wbkSN As String, ar, n As Integer 'Step 1 - Choose Folder to save file
sFolder = Sheets("Dashboard").Range("C20").Value & "\"
Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
        With FolderPicker
            .Title = "Please Choose One"
            .InitialFileName = sFolder
            .AllowMultiSelect = False
            .ButtonName = "Confirm"
                If .Show = -1 Then
                    mypath = .SelectedItems(1)
                    Else
                        Exit Sub
                End If
        End With
    Sheets("Dashboard").Range("C20").Value = mypath
    sFolder = Sheets("Dashboard").Range("C20").Value & "\"
    sFile = Dir(sFolder & "*.xlsm*")
' Loop through the chosen folder
    Do While sFile <> ""
        Set wbkS = Workbooks.Open(sFolder & sFile)
        wbkSN = WorksheetFunction.Substitute(wbkS.Name, "T.xlsm", ".xlsm")
        wbkS.SaveAs Filename:=ThisWorkbook.Sheets("Dashboard").Range("C21").Value & "\" & wbkSN, _
        FileFormat:=52, _
        CreateBackup:=False     'Step 2 - Save as excel
        With wbkS
            ar = .LinkSources(1)    'Step 3 - Update all links in formula
            If Not IsEmpty(ar) Then
                For n = 1 To UBound(ar)
                    .ChangeLink Name:=ar(n), _
                        NewName:=.Name, Type:=xlExcelLinks
                Next
            End If
        End With
        wbkS.Close Savechanges:=False
        Set wbkS = Nothing
        wbkSN = vbNullString
        sFile = Dir
    Loop
Application.DisplayAlerts = True
OptimizedMode False
Set FolderPicker = Nothing
Set wbkS = Nothing
mypath = vbNullString: ar = Empty: n = Empty: sFile = vbNullString: wbkSN = vbNullString
Sheets("Dashboard").Activate
endTime = Now(): timetaken = startTime - endTime
[Status].Value = "Success": [Start_Time].Value = startTime: [Time_Taken].Value = Format(timetaken, "HH:MM:SS"): [UserName].Value = Environ("UserName")
Call captureendtime
MsgBox "Beta template exported to Production"
End Sub
