Attribute VB_Name = "Function_Hardlu"
Option Explicit
Option Compare Text

'=======================================================
Function hardlu(arg As Variant, InputRange As range, outputrange As range) As Variant
'=======================================================
' WHAT IT DOES:
' This function does a lookup even if the input range is
' not nicely sorted, which is what =LOOKUP requires
' Like using VLOOKUP, but without specifying an offset
' Also like using =INDEX/MATCH, but simpler
'=======================================================
Dim i, j As Double
Dim numrows, numcols, numcount As Double
Dim current, lprev, Temp, valrange As Variant

numrows = InputRange.Rows.Count
numcols = InputRange.Columns.Count
If numrows = 1 Then
    numcount = numcols
Else
    numcount = numrows
End If

valrange = InputRange.value
current = ""
    For i = 1 To numcount
         
           If arg = InputRange(i) Then
               current = outputrange(i)
               GoTo finished
            End If
    
    Next i

finished:

hardlu = current
          
End Function



