Attribute VB_Name = "10Recap_Decontruction"
Option Compare Database

Public Function Recap_Decontruction(ByVal Replacement_Factor As Double) As Double
If Replacement_Factor > 0 Then
   Recap_Deconstruction = Replacement_Factor * 0.1
ElseIf Replacement_Factor = 0 Then
    Recap_Deconstruction = "0"
Else
    Replacement_Factor = "Null"
End If
End Function



