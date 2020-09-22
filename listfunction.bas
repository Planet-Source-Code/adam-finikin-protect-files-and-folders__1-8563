Attribute VB_Name = "Module1"
Function ListSave(List As ListBox, FilePath As String)
On Error GoTo error
Dim i As Integer
       On Error GoTo error
       Open FilePath For Output As #1
       For i = 0 To List.ListCount - 1
           Print #1, List.List(i)
       Next i
       Close #1
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function ListOpen(List As ListBox, FilePath As String)
On Error GoTo error
    Dim MyString As String
       On Error GoTo error
       Open FilePath For Input As #1
       While Not EOF(1)
           Input #1, MyString$
           DoEvents
               List.AddItem MyString$
           Wend
           Close #1
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function
