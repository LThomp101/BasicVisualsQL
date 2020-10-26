Attribute VB_Name = "VBA for Class Category"
Option Compare Database


Public Function Class_Category(ByVal MDI As Double, Age As Double) As Variant
If MDI <= 20 And Age <= 30 Then
   Class_Category = "Class A"
ElseIf MDI >= 20 And Age <= 30 Then
    Class_Category = "Class AX"
ElseIf MDI >= 20 And Age >= 30 Then
    Class_Category = "Class AY"
ElseIf MDI > 30 And MDI < 70 And Age > 31 And Age < 50 Then
    Class_Category = "Class B"
ElseIf MDI > 30 And MDI > 70 And Age > 31 And Age < 50 Then
    Class_Category = "Class BX"
ElseIf MDI < 30 And MDI < 70 And Age > 31 And Age < 50 Then
    Class_Category = "Class BY"
ElseIf MDI < 30 And MDI > 70 And Age > 31 And Age < 50 Then
    Class_Category = "BZ"
ElseIf MDI >= 51 And Age >= 71 Then
    Class_Category = "Class C"
ElseIf MDI <= 51 And Age >= 71 Then
    Class_Category = "Class CX"
ElseIf MDI >= 51 And Age <= 71 Then
    Class_Category = "Class CY"
Else
    Class_Category = "Other"
End If
End Function

