' More on how to use Regular Expressions here: http://stackoverflow.com/a/22542835

Option Explicit

Function get_first_match(RE_pattern As String, within_text As String) As String
' Return the first substring of within_text that matches the pattern specified.

    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    RE.pattern = RE_pattern
    RE.Global = False
    
    get_first_match = RE.Execute(within_text).Item(0).Value

End Function
