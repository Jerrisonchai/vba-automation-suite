Attribute VB_Name = "BulkSignature"
' === Module: BulkSignature ===
Option Explicit

' ---- Config (edit these) ----
'Private Const BRAND As String = "Jerrison"
'Private Const REPO As String = "VBA-Utility-Library/vba-file-folder-utils"
'Private Const VERSION_TAG As String = "v2025-08-18"

' Toggle layers
Private Const DO_ABOUT_SHEET As Boolean = True
Private Const DO_CUSTOM_PROP As Boolean = True
Private Const DO_DEFINED_NAME As Boolean = True
Private Const DO_PRINT_FOOTER As Boolean = True
Private Const DO_CODE_MODULE As Boolean = False   ' Requires VBIDE trust & reference

' Office enum fallbacks (avoid extra references)
Private Const msoAutomationSecurityForceDisable As Long = 3
Private Const msoFileDialogFolderPicker As Long = 4
Private Const msoPropertyTypeString As Long = 4

Public Sub SignAllXlsmInFolder()
    Call MyShape_Click
    Call capturetime
    Dim folderPath As String
    folderPath = PickFolder()
    If Len(folderPath) = 0 Then Exit Sub

    Dim backupPath As String
    backupPath = MakeBackupFolder(folderPath)

    Dim oldSec As Long: oldSec = Application.AutomationSecurity
    Dim t0 As Single: t0 = Timer

    Dim cnt As Long, ok As Long, errCnt As Long
    Dim f As String, fullPath As String
    Dim wb As Workbook

'    Dim logS As Worksheet: Set logS = GetOrCreateLog()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    f = Dir(folderPath & "\*.xlsm")
    Do While Len(f) > 0
        fullPath = folderPath & "\" & f
        On Error Resume Next
        FileCopy fullPath, backupPath & "\" & f  ' backup first
        On Error GoTo 0

        On Error GoTo FailOpen
        Set wb = Application.Workbooks.Open(Filename:=fullPath, ReadOnly:=False, UpdateLinks:=0)
        cnt = cnt + 1
        wb.Sheets(1).Activate
'        wb.Worksheets("Dashboard").Unprotect Password:="admin"
        ApplySignature wb
'        wb.Worksheets("Dashboard").Protect Password:="admin"
        wb.Save
        wb.Close SaveChanges:=False
        ok = ok + 1
        'LogRow logS, fullPath, "OK", ""
        GoTo NextFile

FailOpen:
        errCnt = errCnt + 1
        If Not wb Is Nothing Then On Error Resume Next: wb.Close False: On Error GoTo 0
'        LogRow logS, fullPath, "ERROR", Err.Description
        Err.Clear

NextFile:
        f = Dir
    Loop

Cleanup:
    Application.AutomationSecurity = oldSec
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Call captureendtime
    Dim msg As String
    msg = "Processed: " & cnt & vbCrLf & "OK: " & ok & vbCrLf & "Errors: " & errCnt & _
          vbCrLf & "Time (s): " & Format$(Timer - t0, "0.0")
    MsgBox msg, vbInformation, "Bulk Signature"
End Sub

' --- Signature orchestration ---
Private Sub ApplySignature(ByVal wb As Workbook)
    Dim guid As String, sigText As String
    guid = MakeGuid()
    Dim brand As String, repo As String, VERSION_TAG As String
    brand = GetBrand()
    repo = GetRepo()
    VERSION_TAG = GetVersionTag()

    sigText = "Template produced by " & brand & " | Repo: " & repo & _
              " | " & VERSION_TAG & " | SignedOn=" & Format$(Now, "yyyy-mm-dd hh:nn") & _
              " | GUID=" & guid

    If DO_CUSTOM_PROP Then ApplyCustomProperty wb, "JerrisonSignature", sigText
    If DO_DEFINED_NAME Then ApplyDefinedName wb, "_JERR", "=""Template produced by " & brand & """"
    If DO_PRINT_FOOTER Then ApplyPrintFooter wb, "Template produced by " & brand
End Sub

' --- Layers ---
Private Sub ApplyCustomProperty(ByVal wb As Workbook, ByVal propName As String, ByVal propValue As String)
    wb.BuiltinDocumentProperties("Author").Value = propName
    wb.BuiltinDocumentProperties("Comments").Value = propValue
End Sub

Private Sub ApplyDefinedName(ByVal wb As Workbook, ByVal nm As String, ByVal refersTo As String)
    On Error Resume Next
    wb.Names(nm).Delete
    On Error GoTo 0
    wb.Names.Add Name:=nm, refersTo:=refersTo, Visible:=False
End Sub

Private Sub ApplyPrintFooter(ByVal wb As Workbook, ByVal footerText As String)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        On Error Resume Next
        With ws.PageSetup
            .CenterFooter = "&""Calibri,Regular""&8" & footerText
        End With
        On Error GoTo 0
    Next ws
End Sub

' --- Helpers ---

' Instead of fixed Const, we’ll read from Dashboard sheet
' Dashboard sheet must exist in the workbook that runs this macro (Signer.xlsm)
Private Function GetBrand() As String
    On Error Resume Next
    GetBrand = ThisWorkbook.Sheets("Dashboard").Range("C6").Value
    If Len(Trim(GetBrand)) = 0 Then GetBrand = "Jerrison"  ' fallback default
    On Error GoTo 0
End Function

Private Function GetRepo() As String
    On Error Resume Next
    GetRepo = ThisWorkbook.Sheets("Dashboard").Range("C8").Value
    If Len(Trim(GetRepo)) = 0 Then GetRepo = "VBA-Project"
    On Error GoTo 0
End Function

Private Function GetVersionTag() As String
    On Error Resume Next
    GetVersionTag = ThisWorkbook.Sheets("Dashboard").Range("C10").Value
    If Len(Trim(GetVersionTag)) = 0 Then GetVersionTag = "v2025-MM-DD"
    On Error GoTo 0
End Function

Private Function MakeGuid() As String
    On Error Resume Next
    MakeGuid = CreateObject("Scriptlet.TypeLib").guid  ' e.g., "{xxxxxxxx-xxxx-...}"
    If Len(MakeGuid) = 0 Then
        Randomize
        MakeGuid = "RND-" & Format$(Now, "yyyymmddhhmmss") & "-" & Format$(Rnd * 1000000, "000000")
    End If
End Function

Private Function PickFolder() As String
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select folder containing .xlsm files"
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

Private Function MakeBackupFolder(ByVal baseFolder As String) As String
    Dim ts As String
    ts = Format$(Now, "yyyymmdd_HHmmss")
    MakeBackupFolder = baseFolder & "\backup_" & ts
    On Error Resume Next
    MkDir MakeBackupFolder
    On Error GoTo 0
End Function

'Private Function GetOrCreateLog() As Worksheet
'    Dim sh As Worksheet
'    On Error Resume Next
'    Set sh = ThisWorkbook.Worksheets("SignLog")
'    On Error GoTo 0
'    If sh Is Nothing Then
'        Set sh = ThisWorkbook.Worksheets.Add
'        sh.Name = "SignLog"
'        sh.Range("A1:D1").Value = Array("File", "Status", "Details", "Timestamp")
'        sh.Rows(1).Font.Bold = True
'    End If
'    Set GetOrCreateLog = sh
'End Function

Private Sub LogRow(ByVal sh As Worksheet, ByVal filePath As String, ByVal status As String, ByVal details As String)
    With sh
        Dim r As Long: r = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(r, 1).Value = filePath
        .Cells(r, 2).Value = status
        .Cells(r, 3).Value = details
        .Cells(r, 4).Value = Now
    End With
End Sub


