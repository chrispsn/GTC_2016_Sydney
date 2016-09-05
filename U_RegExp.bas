Option Explicit

Public Function get_first_match(RE_pattern As String, within_text As String) As String
' Return the first substring of within_text that matches the pattern specified.

    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    RE.pattern = RE_pattern
    RE.Global = False
    
    get_text = RE.Execute(within_text).Item(0).Value

End Function
