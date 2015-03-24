Attribute VB_Name = "Function_Member"
Option Explicit
Option Compare Text


Function nonmember(arg As Variant, items As Object, Optional lencount As Integer) As Variant
'=================================
' WHAT THIS FUNCTION DOES:
' This function checks to see if an argument is NOT part of a given range.
' If the argument is NOT part of the given range, it returns a 1 , otherwise it returns a 0
' This is done by computing 1 minus the result of the MEMBER function
'
' Also, the optional third argument allows for a match based on the first x characters.
'
' Example: Let's say we want to check if "Apple" in A1 is part of this list in B1:B4: Banana, Cherry, Watermelon, and Strawberry.
' You would write =Nonmember(A1, B1:B4) and it would return a 1, because Apple is not part of that list.
'
' By Jason Chroman
' github.com/jaschro
'=================================
nonmember = 1 - member(arg, items, lencount)
End Function

Function member(arg As Variant, items As Object, Optional lencount As Integer) As Variant
'=================================
' WHAT THIS FUNCTION DOES:
' This function checks to see if an argument is part of a given range.
' If it is, it returns a 1 , otherwise it returns a 0
' This is similar to what the MATCH does, but simpler.
'
' Also, the optional third argument allows for a match based on the first x characters.
'
' Example: Let's say we want to check if "Apple" in A1 is part of this list in B1:B4: Banana, Cherry, Watermelon, and Strawberry.
' You would write =Member(A1, B1:B4) and it would return a 0, because Apple is not part of that list.
'
' By Jason Chroman
' github.com/jaschro
'=================================
Dim i, num As Integer
Dim leftarg, leftitem As String

num = items.count
member = 0

If lencount > 0 Then
    leftarg = Left(arg, lencount)
    For i = 1 To num
        leftitem = Left(items(i), lencount)
        If leftarg = leftitem Then
            member = 1
            GoTo 99
        End If
    Next i
Else
    For i = 1 To num
        If arg = items(i) Then
            member = 1
            GoTo 99
        End If
    Next i
End If

'END =============================
99

End Function













Function lu(arg As Variant, inputrange As range) As Variant
'=======================================================
' WHAT IT DOES:
' This is a simple little lookup function does a lookup
' even if the input range is not nicely sorted, which is
' what =LOOKUP and =VLOOKUP require
' - Unlike the =lookup function, you only have to specify 2 ranges, not 3.
' - It works both in vertical and horizontal modes
' - Since it uses brute force, it's not terribly efficient, but it's plenty fast
'    for small data sets.
'=======================================================
Dim i, j, k, maxloop As Integer
Dim stepcount As Integer
Dim numrows As Integer
Dim numcols As Integer
Dim current As Variant

numrows = inputrange.Rows.count
numcols = inputrange.Columns.count

If numcols = 2 Then
    maxloop = numrows * 2
    k = 1
    stepcount = 2
    ElseIf numrows = 2 Then
        maxloop = numcols * 2
        k = numcols
        stepcount = 1
    Else
        End
End If

current = ""
    
    For i = 1 To maxloop Step stepcount
            j = inputrange(i)
            If arg = j Then
               current = inputrange(i + k)
               GoTo finished
            End If
    
    Next i

finished:

lu = current

End Function

Function efftaxrate(income As Double, inc_range As Object, rates As Object) As Double
    efftaxrate = tax(income, inc_range, rates) / income
End Function


Function tax(income As Double, inc_range As Object, rates As Object) As Double
'=======================================================
' WHAT IT DOES:
' This function, given an income and a tax rate schedule, calculates the tax due.
'
'=======================================================

Dim w, x, ci, cr As Integer
Dim j, k, m, irw  As Double
Dim prodrange As Double

ci = inc_range.count
cr = rates.count
tax = 0

If ci <> cr Then
    GoTo 99
End If

For x = 1 To ci
    If x = 1 Then
        prodrange = Application.WorksheetFunction.Min(income, inc_range(1))
    ElseIf x < ci Then
        j = inc_range(x) - inc_range(x - 1)
        k = income - inc_range(x - 1)
        m = Application.WorksheetFunction.Min(k, j)
        prodrange = Application.WorksheetFunction.Max(m, 0)
    ElseIf x = ci Then
        prodrange = Application.WorksheetFunction.Max(income - inc_range(x - 1), 0)
    End If
    tax = tax + prodrange * rates(x)
Next
99
End Function

'=======================================================
Function exclude(inputrange As range, notinrange As range) As Variant
'=======================================================
' WHAT IT DOES:
' This function takes the first item of a list that is not
' part of another list
'=======================================================
Dim i As Integer
Dim j As Integer
Dim numrows As Integer
Dim numcols As Integer
Dim valrange As Variant
Dim temp As Variant
Dim lprev As Variant

numrows = inputrange.Rows.count
valrange = inputrange.Value
current = ""
    For i = 1 To numrows
         
           If member(inputrange(i), notinrange) Then
            Else
               current = inputrange(i)
               GoTo finished
            End If
    
    Next i

finished:

exclude = current
          
End Function



'=======================================================
'=======================================================
Function hardlu(arg As Variant, inputrange As range, outputrange As range) As Variant
'=======================================================
' WHAT IT DOES:
' This function does a lookup even if the input range is
' not nicely sorted, which is what =LOOKUP and =VLOOKUP require
'=======================================================
Dim i As Integer
Dim j As Integer
Dim numrows As Integer
Dim numcols As Integer
Dim valrange As Variant
Dim temp As Variant
Dim lprev As Variant

numrows = inputrange.Rows.count
valrange = inputrange.Value
current = ""
    For i = 1 To numrows
         
           If arg = inputrange(i) Then
               current = outputrange(i)
               GoTo finished
            End If
    
    Next i

finished:

hardlu = current
          
End Function



'=======================================================
'=======================================================

