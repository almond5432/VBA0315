Attribute VB_Name = "Module1"
Option Explicit

Sub demoad()
Dim userstr As String
userstr = InputBox("�п�J�n�Ұʤu�@��W��")

Dim wsheet As Worksheet
For Each wsheet In Worksheets
If (wsheet.Name = userstr) Then
MsgBox "���u�@��G" & wsheet.Name
wsheet.Activate
Else
MsgBox "Sorry,�A�|�����"
End If
Next
MsgBox "�����j�鱽�y"
End Sub
