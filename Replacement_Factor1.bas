Attribute VB_Name = "Replacement_Factor1"
Option Compare Database
Public Function Replacement_Factor(ByVal RPV As Variant, Age As Variant) As Double
If IsNumeric(Age) = False Or IsNumeric(RPV) = False Then
    Replacement_Factor = 0
ElseIf Age >= 60 Then
   Replacement_Factor = [RPV]
Else
    Replacement_Factor = 0
End If
End Function

