Attribute VB_Name = "Module1"
Option Explicit

Sub DEMOmsg()
Msg "maintext", 3, "我是標題"

End Sub

Sub IntDemo()
    Dim i As Integer
    i = 1000
    MsgBox i
    
End Sub

Sub StringDemo2()
Dim i As String
    i = "葉鮭魚想吃壽司郎"
    MsgBox i
End Sub

Sub SingleDemo()
    Dim i As Single
    i = 12.3456
    MsgBox i
End Sub

Sub DoubleDemo()
    Dim i As Double
    i = 12.3456543168796
    MsgBox i
End Sub

Sub boolDemo()
    Dim i As Boolean
    i = True
    MsgBox i
End Sub

Sub DateDemo()
    Dim i As Single
    currentDate = Now
    MsgBox i
End Sub

Sub Quit()
    Dim rst As Integer
    rst = MsgBox("要結束程式嗎？", vbYesNo, "結束程式")
    
    MsgBox i
End Sub

Sub inputBoxDemo()
Dim userString1 As String '宣告變數-宣告一個文字型態變數,名稱叫userString1
'
userString1 = InputBox("你好! 請問你的名字是？")
MsgBox "嗨!" & userString1
Dim userString2 As String
userString2 = InputBox("你喜歡聽什麼音樂？")
MsgBox "哇!原來你喜歡" & userString2
End Sub
