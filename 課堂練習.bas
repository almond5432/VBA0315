Attribute VB_Name = "Module1"
Option Explicit

Sub DEMOmsg()
Msg "maintext", 3, "�ڬO���D"

End Sub

Sub IntDemo()
    Dim i As Integer
    i = 1000
    MsgBox i
    
End Sub

Sub StringDemo2()
Dim i As String
    i = "���D���Q�Y�إq��"
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
    rst = MsgBox("�n�����{���ܡH", vbYesNo, "�����{��")
    
    MsgBox i
End Sub

Sub inputBoxDemo()
Dim userString1 As String '�ŧi�ܼ�-�ŧi�@�Ӥ�r���A�ܼ�,�W�٥suserString1
'
userString1 = InputBox("�A�n! �аݧA���W�r�O�H")
MsgBox "��!" & userString1
Dim userString2 As String
userString2 = InputBox("�A���wť���򭵼֡H")
MsgBox "�z!��ӧA���w" & userString2
End Sub
