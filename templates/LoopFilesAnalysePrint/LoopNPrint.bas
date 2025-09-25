Attribute VB_Name = "LoopNPrint"
Option Explicit

Sub PrintAnalyticsofFiles()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()

Dim folderPath As String
Dim filename As String
Dim wb As Workbook
Dim data_lastrow As Long

'Get folder name, add \ if not available
folderPath = ThisWorkbook.Sheets("Dashboard").Range("C16").Value
If Right(folderPath, 1) <> "\" Then folderPath = folderPath + "\"

filename = Dir(folderPath & "*.xls")

'OptimizedMode True

'start of loop, to open each file in the named folder
'Replace raw data, formulas in Print will calculate, get filename into Print, and print out UsedRange
Do While filename <> ""

'    Application.AskToUpdateLinks = False
'    Application.DisplayAlerts = False
    
    ThisWorkbook.Sheets("RawData").Range("A:AZ").Clear
    
    Set wb = Workbooks.Open(folderPath & filename)
    data_lastrow = Application.CountA(wb.Sheets("Sheet1").Range("A:A"))
    
    ThisWorkbook.Sheets("RawData").Range("A1:AZ" & data_lastrow).Value = wb.Sheets("Sheet1").Range("A1:AZ" & data_lastrow).Value
    ThisWorkbook.Sheets("Print").Range("A1").Value = wb.Name
    ThisWorkbook.Sheets("Dashboard").Range("C10").Value = wb.Name
    
    ThisWorkbook.Sheets("Print").Activate
    Call Application.CalculateFull
    DoEvents
    ThisWorkbook.Sheets("Print").UsedRange.PrintOut     'Make default print setting as PDF to produce multiple pdf copies quickly
    
    wb.Close Savechanges:=False
    
'    Application.DisplayAlerts = True
'    Application.AskToUpdateLinks = True
    
    filename = Dir
Loop

'Dim LastRow As Long
'LastRow = Sheets("Print").Cells(Cells.Rows.Count, "J").End(xlUp).Row + 1
'Sheets("Print").Range("J" & LastRow).Formula = "=SUM(J2:J" & LastRow - 1 & ")"

'    OptimizedMode False
    Set wb = Nothing
    data_lastrow = Empty: filename = vbNullString: folderPath = vbNullString
endTime = Now(): timetaken = startTime - endTime
[Status].Value = "Success": [Start_Time].Value = startTime: [Time_Taken].Value = Format(timetaken, "HH:MM:SS"): [UserName].Value = Environ("UserName")
Call captureendtime
MsgBox "Print done"

End Sub
