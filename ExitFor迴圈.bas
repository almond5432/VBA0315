Attribute VB_Name = "Module1"
Option Explicit

Sub exitfordemo1()
Dim i As Integer
Dim sum As Integer
sum = 0
For i = 1 To 10
    MsgBox "ヘ玡i=" & i '陪ボヘ玡Ω计
    
    sum = sum + i
    MsgBox "ヘ玡羆㎝=" & sum
    
    If i >= 4 Then
    MsgBox "才单4"
    End If


    
    Next
MsgBox "羆㎝ =" & sum

End Sub
Sub exitfordemo2()

Dim i As Integer
Dim sum As Integer
sum = 0
For i = 1 To 10
    MsgBox "ヘ玡i=" & i '陪ボヘ玡Ω计
    
    sum = sum + i
    MsgBox "ヘ玡羆㎝=" & sum
    
    If i >= 4 Then
    MsgBox "才单4"
    Exit For
    
    End If


    
    Next
MsgBox "羆㎝ =" & sum

End Sub


