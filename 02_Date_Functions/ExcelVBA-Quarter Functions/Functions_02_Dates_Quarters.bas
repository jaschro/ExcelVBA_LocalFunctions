Attribute VB_Name = "Functions_02_Dates_Quarters"
Option Explicit

Function quarter(inDate As Variant) As Variant
'==============================================================================
' This function calculates the quarter in which any given date falls.
' For example:
' 1 = January, February, March
' 2 = April, May, June
' 3 = July, August, September
' 4 = October, November, December
'==============================================================================

Dim i As Double
Dim x As Integer

If VBA.isdate(inDate) Then
    i = Month(inDate)
    Select Case i
        Case Is <= 3
            quarter = 1
        Case Is <= 6
            quarter = 2
        Case Is <= 9
            quarter = 3
        Case Is >= 10
            quarter = 4
    End Select
Else
    quarter = ""
End If

End Function


Function qy(inDate As Variant) As String
'==============================================================================
' Converts a date of 5/21/2010 into 2010.2 for easy aggregation of data by quarter
'==============================================================================
If VBA.isdate(inDate) Then
    qy = "Q" & quarter(inDate) & " '" & Right(Year(inDate), 2)
End If

End Function




Public Function yrq(inDate As Variant) As String
'==============================================================================
' Converts a date of 5/21/2010 into 2010.2 for easy aggregation of data by quarter
'==============================================================================
If VBA.isdate(inDate) Then
yrq = Format(Year(inDate) + quarter(inDate) / 10, "###0.0")
End If
End Function

