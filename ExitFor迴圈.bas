Attribute VB_Name = "Module1"
Option Explicit

Sub exitfordemo1()
Dim i As Integer
Dim sum As Integer
sum = 0
For i = 1 To 10
    MsgBox "�ثei=" & i '��ܥثe����
    
    sum = sum + i
    MsgBox "�ثe�`�M=" & sum
    
    If i >= 4 Then
    MsgBox "�ŦX�j�󵥩�4"
    End If


    
    Next
MsgBox "�`�M =" & sum

End Sub
Sub exitfordemo2()

Dim i As Integer
Dim sum As Integer
sum = 0
For i = 1 To 10
    MsgBox "�ثei=" & i '��ܥثe����
    
    sum = sum + i
    MsgBox "�ثe�`�M=" & sum
    
    If i >= 4 Then
    MsgBox "�ŦX�j�󵥩�4"
    Exit For
    
    End If


    
    Next
MsgBox "�`�M =" & sum

End Sub


