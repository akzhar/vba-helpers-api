Attribute VB_Name = "VbaHelper_AddLinkToWs"
Option Explicit

Function AddLinkToWs(ByVal linkRng As Range, ByVal wsName$)
    ' Puts a link to the sheet in specified cell
    linkRng.Parent.Hyperlinks.Add _
        Anchor:=linkRng, _
        Address:="", _
        SubAddress:="'" & wsName & "'!A1", _
        TextToDisplay:=wsName                                    
End Function