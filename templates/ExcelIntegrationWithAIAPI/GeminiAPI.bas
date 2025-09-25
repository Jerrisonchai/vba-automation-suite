Attribute VB_Name = "GeminiAPI"
Option Explicit

Function GetAISummary_Gemini(prompt As String) As String
    Dim http As Object
    Dim JSON As Object
    Dim postData As String
    Dim result As String
    Dim apiKey As String
    
    apiKey = Sheets("Reference").Range("C5").Value   ' << replace
    
    ' Prepare request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" & apiKey, False
    http.setRequestHeader "Content-Type", "application/json"
    
    postData = "{""contents"":[{""parts"":[{""text"":""" & prompt & """}]}]}"
    
    http.Send postData
    result = http.responseText
    
    ' Parse JSON (requires VBA JSON library)
    Set JSON = JsonConverter.ParseJson(result)
    
    On Error Resume Next
    GetAISummary_Gemini = JSON("candidates")(1)("content")("parts")(1)("text")
    If GetAISummary_Gemini = "" Then
        GetAISummary_Gemini = "? Gemini API error: " & result
    End If
    On Error GoTo 0
End Function


'--- Main procedure to log daily tasks & AI comment
Sub LogDailyTasks()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim taskList As String
    Dim aiSummary As String
    
    Set ws = ThisWorkbook.Sheets("Tasks")
    
    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    ' Build today's task string (assumes tasks are in column C starting row 2)
    taskList = ws.Cells(lastRow, "G").Value

    ' Call AI to summarize
    aiSummary = GetAISummary_Gemini("Here are today's tasks: " & vbCrLf & taskList & vbCrLf & "Please summarize into insights for reporting.")
    
    ' Write row
    ws.Cells(lastRow, 8).Value = aiSummary                       ' AI Comment
End Sub


