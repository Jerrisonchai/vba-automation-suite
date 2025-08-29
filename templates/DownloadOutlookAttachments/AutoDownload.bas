Attribute VB_Name = "AutoDownload"
Option Explicit

'app --> name space --> inbox --> items --> Mailitems --> Attachment

Sub download_attachments()
'Call FilenMyShape.MyShape_Click
Dim olApp As Outlook.Application
Dim olNS As Outlook.Namespace
Dim olFolder As Outlook.MAPIFolder
Dim olItem As Object
Dim mailitem As Outlook.mailitem
Dim olAtt As Outlook.Attachment

Dim startTime As Date
Dim endTime As Date
Dim UserName As String
Dim timetaken As Date

startTime = Now()
'On Error GoTo ErrX

Call capturetime
Call MyShape_Click

Set olApp = New Outlook.Application
Set olNS = olApp.GetNamespace("MAPI")

Set olFolder = olNS.Folders([Mailbox_Name].Text)
Set olFolder = olFolder.Folders("Inbox")
Set olFolder = olFolder.Folders("Purchasing Project")
Set olFolder = olFolder.Folders("from Supplier")


For Each olItem In olFolder.Items

    If olItem.Class = olMail Then
        Set mailitem = olItem
        
        Debug.Print mailitem.Subject
        Debug.Print mailitem.ReceivedTime
        
       
        For Each olAtt In mailitem.Attachments
            olAtt.SaveAsFile [Export_To].Text & "\" & olAtt.Filename
        Next olAtt
    
    End If
Next olItem

Set olApp = Nothing
Set olNS = Nothing
Set olFolder = Nothing
Set olItem = Nothing
Set mailitem = Nothing
Set olAtt = Nothing

endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Success"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
'[User_Name].Value = Environ("UserName")
MsgBox "Download done!"
Exit Sub
ErrX:
endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Failed"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
[User_Name].Value = Environ("UserName")

Call captureendtime
MsgBox "Download done!"
End Sub

Sub SaveOlAttachments()
Call capturetime
    Dim olFolder As Outlook.MAPIFolder
    Dim msg As Outlook.mailitem
    Dim msg2 As Outlook.mailitem
    Dim att As Outlook.Attachment
    Dim strFilePath As String
    Dim strTmpMsg As String
    Dim fsSaveFolder As String
    
    Dim Reg1 As RegExp
    Dim M1 As MatchCollection
    Dim M As Match
    Dim strURL As String

    fsSaveFolder = [Export_To].Text & "\"

    'path for creating attachment msg file for stripping
    strFilePath = [Export_To].Text & "\"
    strTmpMsg = "KillMe.msg"
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.Namespace
    Set olApp = New Outlook.Application
    Set olNS = olApp.GetNamespace("MAPI")
   'Get outlook folder from Dashboard
    Set olFolder = olNS.Folders([Mailbox_Name].Text)
    Set olFolder = olFolder.Folders(Sheets("Dashboard").Range("C18").Value)
    Set olFolder = olFolder.Folders(Sheets("Dashboard").Range("D18").Value)
    Set olFolder = olFolder.Folders(Sheets("Dashboard").Range("E18").Value)
    If olFolder Is Nothing Then Exit Sub

    For Each msg In olFolder.Items
        If msg.Attachments.Count > 0 Then
            Dim A As Long
            A = 1
                Do Until A = msg.Attachments.Count + 1
                        Dim bflag
                        bflag = False
                        If Right$(msg.Attachments(A).Filename, 3) = "msg" Then
                            bflag = True
                            msg.Attachments(A).SaveAsFile strFilePath & strTmpMsg
                            Set msg2 = Outlook.CreateItemFromTemplate(strFilePath & strTmpMsg)
                        End If
                        On Error Resume Next
                        If bflag Then
                            Dim sSavePathFS
                            sSavePathFS = fsSaveFolder & msg2.Attachments(A).Filename
                            msg2.Attachments(A).SaveAsFile sSavePathFS
                            msg2.Delete
                        Else
                            sSavePathFS = fsSaveFolder & msg.Attachments(A).Filename
                            msg.Attachments(A).SaveAsFile sSavePathFS
                        End If
                    A = A + 1
                Loop
        Else
        
            Set Reg1 = New RegExp
            With Reg1
                .Pattern = "View catalog <(.*)>"
                .Global = True
                .IgnoreCase = True
            End With
            If Reg1.Test(msg.Body) Then
                Set M1 = Reg1.Execute(msg.Body)
                For Each M In M1
                    strURL = M.SubMatches(0)
                    Debug.Print strURL
                        Shell ("C:\Program Files\Google\Chrome\Application\chrome.exe -url " & strURL)
NextURL:
                Next
            End If
            
        End If
    Next
    
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder([Export_To].Text)
    For Each oFile In oFolder.Files
        If oFile.Name Like "*.xls*" Then
        Else
            oFile.Delete
        End If
    Next
    
Set olFolder = Nothing
Set msg = Nothing
Set msg2 = Nothing
Set att = Nothing
Set Reg1 = Nothing
Set olApp = Nothing
Set olNS = Nothing
Set oFSO = Nothing
Set oFolder = Nothing
Set oFile = Nothing

Call captureendtime
MsgBox "Download done!"

End Sub

Public Sub OpenLinksMessage(olMail As Outlook.mailitem)

Dim Reg1 As RegExp
Dim M1 As MatchCollection
Dim M As Match
Dim strURL As String
Dim oApp As Object
Set oApp = CreateObject("InternetExplorer.Application")

Set Reg1 = New RegExp

With Reg1
.Pattern = "(https?[:]//([0-9a-z=\?:/\.&-^!#$%;_])*)"
.Global = True
.IgnoreCase = True
End With

''Wait for a certain amount of time before opening URLs.
'tTime0 = Now
'Do Until tTime0 + TimeValue("00:00:53") < Now
'DoEvents
'Loop

If Reg1.Test(olMail.Body) Then

    Set M1 = Reg1.Execute(olMail.Body)
    For Each M In M1
        strURL = M.SubMatches(0)
        Debug.Print strURL
        If InStr(strURL, "@") Then GoTo NextURL 'Ignore emails.
        If InStr(strURL, "unsubscribe") Then GoTo NextURL 'Ignore specific string in URL.
        If Right(strURL, 1) = ">" Then strURL = Left(strURL, Len(strURL) - 1)
        oApp.navigate strURL, CLng(2048)
        oApp.Visible = True
        
        'wait for page to load before passing the web URL
        Do While oApp.Busy
        DoEvents
        Loop
    
NextURL:
    Next
    
End If

Set Reg1 = Nothing
End Sub


Sub clicklinks()

Dim olApp As Outlook.Application
Dim olNS As Outlook.Namespace
Dim olFolder As Outlook.MAPIFolder
Dim olItem As Object
Dim mailitem As Outlook.mailitem

Set olApp = New Outlook.Application
Set olNS = olApp.GetNamespace("MAPI")

Set olFolder = olNS.Folders([Mailbox_Name].Text)
    Set olFolder = olFolder.Folders(Sheets("Dashboard").Range("C18").Value)
    Set olFolder = olFolder.Folders(Sheets("Dashboard").Range("D18").Value)
    Set olFolder = olFolder.Folders(Sheets("Dashboard").Range("E18").Value)

Dim Reg1 As RegExp
Dim M1 As MatchCollection
Dim M As Match
Dim strURL As String

For Each olItem In olFolder.Items

    If olItem.Class = olMail Then

        Set mailitem = olItem
        Set Reg1 = New RegExp

    With Reg1
'        .Pattern = "(https?[:]//([0-9a-z=\?:/\.&-^!#$%;_])*)"
        .Pattern = "View catalog <(.*)>"
        .Global = True
        .IgnoreCase = True
    End With

    If Reg1.Test(mailitem.Body) Then
            Set M1 = Reg1.Execute(mailitem.Body)

            For Each M In M1
                strURL = M.SubMatches(0)
                Debug.Print strURL
                    Shell ("C:\Program Files\Google\Chrome\Application\chrome.exe -url " & strURL)
NextURL:
            Next
    End If

    End If
Next olItem

Set olApp = Nothing
Set olNS = Nothing
Set olFolder = Nothing
Set olItem = Nothing
Set mailitem = Nothing

Set Reg1 = Nothing

MsgBox "Done!"
End Sub
