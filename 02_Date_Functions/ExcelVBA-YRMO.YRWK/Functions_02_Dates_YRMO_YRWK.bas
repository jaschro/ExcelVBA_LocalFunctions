Attribute VB_Name = "Functions_02_Dates_YRMO_YRWK"
Option Explicit

Public Function yrmo(baseDate As Variant) As String
'==============================================================================
' Converts a date of 5/21/2010 into 2010.05 for easy aggregation of data by month
'==============================================================================
'yrmo = Year(baseDate) + Month(baseDate) / 100
If isdate(baseDate) Then
    yrmo = Format(Year(baseDate) + Month(baseDate) / 100, "###0.00")
Else
    yrmo = ""
End If
End Function


Public Function yrwk(baseDate As Variant) As String
'==============================================================================
' Converts a date of 5/21/2010 into 2010.21 for easy aggregation of data by year and week
'==============================================================================
Dim week As Variant
If isdate(baseDate) Then
    week = Application.WorksheetFunction.WeekNum(baseDate)
    yrwk = Format(Year(baseDate) + week / 100, "###0.00")
Else
    yrwk = ""
End If
End Function


