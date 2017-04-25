Set regexNumber = New RegExp

regexNumber.Global = True
regexNumber.IgnoreCase = True

regexNumber.Pattern = "^\d+$"

Dim inputString
inputString = InputBox("Computer name?")

If regexNumber.Test( inputString ) Then
    WScript.Echo "True' Number"
Elseif inputString = "Tom" Then
    WScript.Echo "False ' It's invalid"
End If
