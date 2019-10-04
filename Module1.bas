Option Explicit
'---Button A---
Sub DraftEmails()
Dim Class As New Class1
Set Class = New Class1
With Class
    .Response = MsgBox("Are you sure you want to save the emails to your draft folder?", vbYesNo, "Drafs?")
        If .Response <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
        
    'Emails---------------------
    .EmailPath = .CurrentPath & "\" & "Template.msg"
    If Dir(.EmailPath) = "" Then
        MsgBox "Could not find the file " & .EmailPath
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 3 To .LastRow
        .Msg(.EmailPath, i).Save
    Next i
    
    Application.ScreenUpdating = True
    MsgBox ("Email(s) successfully drafted")
End With
Set Class = Nothing
End Sub

'---Button B---
Sub SendEmails()
Dim Class As New Class1
Set Class = New Class1
With Class
    .Response = MsgBox("Do you want to send email(s) now?", vbYesNo, "Send?")
        If .Response <> vbYes Then Exit Sub
    .Response = MsgBox("Are you sure you want to continue?", vbYesNo, "Send?")
        If .Response <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    
    'Emails---------------------
    Dim i As Integer
    For i = 3 To .LastRow
        .Msg(.Filepath & "\" & "Template.msg", i).Display
        SendKeys "%{S}"
    Next i
    
    Application.ScreenUpdating = True
    MsgBox ("Email(s) successfully sent")
End With
Set Class = Nothing
End Sub

'---Button C---
Sub ClearALL()
Dim Response As String

Response = MsgBox("Are you sure you want to clear everything", vbYesNo, "Clear Data")
    If Response <> vbYes Then Exit Sub
    Rows("3:" & Rows.Count).Delete Shift:=xlUp
    Range("A1").Select
End Sub
