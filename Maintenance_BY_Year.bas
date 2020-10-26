Attribute VB_Name = "Maintenance_BY_Year"
Option Compare Database

Public Function Maintenance_Yearly(ByVal Overall_Asset_Condition As Variant, Repair_Needs As Variant) As Double
If IsNumeric(Repair_Needs) = False Or IsNull(Overall_Asset_Condition) = True Then
    Maintenance_Yearly = 0
ElseIf Overall_Asset_Condition = "Adequate" Then
    Maintenance_Yearly = Repair_Needs * 0.0005
ElseIf Overall_Asset_Condition = "Substandard" Then
    Maintenance_Yearly = Repair_Needs * 0.001
ElseIf Overall_Asset_Condition = "Inadequate" Then
    Maintenance_Yearly = Repair_Needs * 0.015
Else
    Maintenance_Yearly = 0
End If
End Function
