Option Explicit

' Function to demonstrate safe type handling
Function AddSafe(a, b)
  If IsNumeric(a) And IsNumeric(b) Then
    AddSafe = CDbl(a) + CDbl(b)
  Else
    AddSafe = "Error: Non-numeric input"
  End If
End Function

' Example usage
Dim result
result = AddSafe(5, 10) ' Correct: Returns 15
MsgBox result

result = AddSafe("5", 10) ' Correctly handles string input
MsgBox result

result = AddSafe("abc", 10) ' Correctly handles non-numeric string
MsgBox result

' Early Binding Example (if applicable to the specific bug)
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")  'Early binding for better error checking
' ... use objFSO ...
Set objFSO = Nothing