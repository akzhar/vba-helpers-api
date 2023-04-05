Attribute VB_Name = "Helper98"
Option Explicit

Function GetRegExpSubMatches(ByVal text$, ByVal pattern$, Optional ByVal keepFirstSub As Boolean = True) As Variant()
    ' Gets all regular expression sub matches from the text

    GetRegExpSubMatches = Array()

    Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
    regExp.pattern = pattern
    regExp.Global = True
    regExp.MultiLine = True
    
    If regExp.test(text) Then
        Dim matchesColl As Object: Set matchesColl = regExp.Execute(text)
        If matchesColl.Count <> 0 Then
            Dim i&, allMatches()
            For i = 0 To matchesColl.Count - 1
                Dim j&, elem, subMatches()
                For j = 0 To matchesColl(i).submatches.Count - 1
                    elem = matchesColl(i).submatches(j)
                    Call AddToArr(subMatches, elem) ' @dependency: 1.bas
                Next j
                If Not keepFirstSub Then
                  Call AddToArr(allMatches, subMatches) ' @dependency: 1.bas
                  Erase subMatches
                End if
            Next i
            GetRegExpSubMatches = Iif(keepFirstSub, subMatches, allMatches)
        End If
    End If
End Function