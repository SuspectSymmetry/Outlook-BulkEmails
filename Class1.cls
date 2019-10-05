Option Explicit
Public Response, EmailPath As String

Public Function CurrentPath() As String
CurrentPath = Application.ActiveWorkbook.Path
End Function

Public Function LastRow() As Long
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
End Function

Public Function Msg(Email As String, i As Integer) As Object
On Error GoTo Handler:
Dim olApp As Object
Set olApp = CreateObject("Outlook.Application")
Set Msg = olApp.CreateItemFromTemplate(Email)

With Msg
    .SentOnBehalfOfName = Cells(2, 4)
    .To = Cells(i, 1)
    .Cc = Cells(i, 2)
    .BCC = Cells(i, 3)
    .Recipients.ResolveAll
End With
Exit Function

Handler:
    MsgBox Err.Description
End Function
