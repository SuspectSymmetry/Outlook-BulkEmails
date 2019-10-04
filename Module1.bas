Option Explicit

Sub DraftEmails()
Dim olApp As Object
Dim Msg As Object
Dim e As Integer
Dim Response, FilePath, SigString, Signature As String
Dim Lastrow As Long
Dim Environ As Object

Response = MsgBox("Are you sure you want to save the emails to your draft folder?", vbYesNo, "Drafs?")
    If Response <> vbYes Then Exit Sub
Application.ScreenUpdating = False

'Emails
Lastrow = Range("A" & Rows.Count).End(xlUp).Row
FilePath = Application.ActiveWorkbook.Path
Set olApp = CreateObject("Outlook.Application")

    For e = 3 To Lastrow
        Set Msg = olApp.CreateItemFromTemplate(FilePath & "\" & "Template.msg")
            With Msg
            .SentOnBehalfOfName = Cells(2, 4)
            .To = Cells(e, 1)
            .Cc = Cells(e, 2)
            .BCC = Cells(e, 3)
            '.Body = Cells(e, 4)
            .Recipients.ResolveAll
            .Save
            End With
    Next e

Application.ScreenUpdating = True
MsgBox ("Email(s) successfully drafted")
Set olApp = Nothing
Set Msg = Nothing

End Sub

Sub SendEmails()
Dim olApp As Object
Dim Msg As Object
Dim e As Integer
Dim Response, FilePath, SigString, Signature As String
Dim Lastrow As Long
Dim Environ As Object

Response = MsgBox("Do you want to send email(s) now?", vbYesNo, "Send?")
    If Response <> vbYes Then Exit Sub

Response = MsgBox("Are you sure you want to continue?", vbYesNo, "Send?")
    If Response <> vbYes Then Exit Sub
Application.ScreenUpdating = False


Application.ScreenUpdating = False
'On Error Resume Next
'--------------------------------------

'Emails
Lastrow = Range("A" & Rows.Count).End(xlUp).Row
FilePath = Application.ActiveWorkbook.Path
Set olApp = CreateObject("Outlook.Application")

    For e = 3 To Lastrow
        Set Msg = olApp.CreateItemFromTemplate(FilePath & "\" & "Template.msg")
            With Msg
            .SentOnBehalfOfName = Cells(2, 4)
            .To = Cells(e, 1)
            .Cc = Cells(e, 2)
            .BCC = Cells(e, 3)
            '.Body = Cells(e, 4)
            .Recipients.ResolveAll
            .Display
            SendKeys "%{S}"
            End With
    Next e
      
Application.ScreenUpdating = True
MsgBox ("Email(s) successfully sent")
Set olApp = Nothing
Set Msg = Nothing

End Sub

Sub ClearALL()
Dim Response As String

Response = MsgBox("Are you sure you want to clear everything", vbYesNo, "Clear Data")
    If Response <> vbYes Then Exit Sub
    Rows("3:" & Rows.Count).Delete Shift:=xlUp
    Range("A1").Select
End Sub
