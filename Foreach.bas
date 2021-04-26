Attribute VB_Name = "Module1"
Option Explicit

Sub demoad()
Dim userstr As String
userstr = InputBox("請輸入要啟動工作表名稱")

Dim wsheet As Worksheet
For Each wsheet In Worksheets
If (wsheet.Name = userstr) Then
MsgBox "找到工作表：" & wsheet.Name
wsheet.Activate
Else
MsgBox "Sorry,你尚未找到"
End If
Next
MsgBox "完成迴圈掃描"
End Sub
