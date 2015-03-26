Attribute VB_Name = "Functions_03_Text_Concat"
Function concat(avec As Variant, Optional CHAR2INS As String) As String
'========================================================
' This is a simple function that improves on Excel's built in =concatenate function
' Arguments are:
' AVEC - this is the vector to be concatenated
' CHAR2INS - is the charachter to insert between cell items.  If left blank, Nothing Is inserted
'========================================================
Dim i, j, counter, Total As Integer
Dim numrows, numcols As Integer
Dim StrToReturn As String
StrToReturn = ""
numrows = avec.Rows.Count
numcols = avec.Columns.Count
Total = numrows * numcols
For j = 1 To numrows
    For i = 1 To numcols
        ' skip blanks
        If avec(j, i) <> "" Then
        
        ' don't put a character after the last item
            counter = counter + 1
            If counter = Total Then
                StrToReturn = StrToReturn & avec(j, i)
            Else
                StrToReturn = StrToReturn & avec(j, i) & CHAR2INS
            End If
        End If

    Next i
Next j

' this line takes out unprintable characters
concat = Application.trim(StrToReturn)

End Function


