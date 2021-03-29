Attribute VB_Name = "Module2"
Option Explicit

Public Function eoqdemo2(demand, holding, fixed) As Integer
eoqdemo2 = (2 * demand * fixed / holding) ^ (1 / 2)
End Function

Public Function BMI(weight, hight)
BMI = weight / (hight / 100) ^ 2
End Function

