Option Explicit

Function sort_array(some_array)
' Inspired by https://en.m.wikipedia.org/wiki/Windows_Script_File#Mixed_language_support
' TODO still need to figure this one out - intermediate results are being converted to strings?

    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    With sc
        .Language = "JScript"
        .AddCode "function sort_array(VBA_array) {return VBA_array.toArray().sort();}"
        Dim result As Variant
        result = .Run("sort_array", some_array)
    End With
    sort_array = Split(result, ",")

End Function
