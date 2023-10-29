Attribute VB_Name = "VbaHelper_AddHyperLink"
Option Explicit

Function AddHyperLink(ByVal linkRng As Range, ByVal url$, ByVal displayText$)
    ' Puts a hyperlink in specified cell
    linkRng.Parent.Hyperlinks.Add _
        Anchor:=linkRng, _
        Address:=url, _
        TextToDisplay:=displayText
End Function