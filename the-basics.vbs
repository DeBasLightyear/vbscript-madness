' Variables
Dim aVariableWithAString
aVariableWithAString = "Hello There"

Dim aVariableWithANumber
aVariableWithANumber = 1.5

Dim aVariableWithABoolean
aVariableWithABoolean = True
aVariableWithABoolean = False


' Arithmatic operators
result = 1 + 1
result = 1 - 1
result = 1 * 2
result = 1 / 2
result = 1 % 0.5
result = 2 ^ 2

' Comparison operators
result = 1 = 1
result = 1 <> 1
result = 1 > 2
result = 1 < 2
result = 1 >= 0.5
result = 2 <= 2

' Logical operators
result = 1 = 1 AND 2 = 2
result = 1 = 1 OR 2 = 1
result = NOT(1 <> 2)
result = 1 <> 0 XOR 2 <> 1

' Concatenation operator
result = "Hello" & " " & "there"

' Arrays
Dim aNiceArrayName()                ' declareer een array met de haakjes
Dim(0) = "Dit is een waarde"        ' ken een waarde toe aan de eerste positie
Dim(1) = 42                         ' ken een waarde toe aan de tweede positie
anotherVariable = aNiceArrayName(0) ' pak de waarde van de tweede positie

' Decision making - If
If 2 > 1 Then
  ' Doe iets hier
End If

' Decision making - If-Else
If IsNumeric("1") Then
  ' Doe iets
Else
  ' Doe iets anders
End If

' Loops - For Each
For Each thingy In aNiceArray
  ' Do something
Next

' Loops - Do-Loop Until
Do
  ' Do something
Loop Until 42 > 5

' Functions
Function functieNaam(parameter1, parameter2)
  ' Do something
  ' Return something
  functieNaam = [resultaat]
End Function

Dim output
output = functieNaam(1, 2)

' Sub procedures
Sub subNaam(parameter1, parameter2, parameter3)
  ' Doe iets
End Sub

Call subNaam(1, 2, 3)
