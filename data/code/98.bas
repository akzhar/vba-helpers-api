Attribute VB_Name = "Helper98"
Option Explicit

Function GetRegExpSubMatches(ByVal text$, ByVal pattern$) As Variant()
    ' Gets all regular expression sub matches from the text

    GetRegExpSubMatches = Array()

    Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
    regExp.pattern = pattern
    regExp.Global = True
    regExp.MultiLine = True
    
    If regExp.test(text) Then
        Dim matchesColl As Object: Set matchesColl = regExp.Execute(text)
        If matchesColl.Count <> 0 Then
            Dim i&, matches()
            For i = 0 To matchesColl.Count - 1
                Dim j&, arr()
                For j = 0 To matchesColl(i).submatches.Count - 1
                    Call AddToArr(arr, matchesColl(i).submatches(j)) ' @(id 1)
                Next j
                Call AddToArr(matches, arr) ' @(id 1)
                Erase arr
            Next i
            GetRegExpSubMatches = matches
        End If
    End If
End Function