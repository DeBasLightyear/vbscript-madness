' Opdracht 1A: Hello World
' Geef een Message Box weer die het bericht "Hello World" weergeeft (of een ander naar smaak).
MsgBox("Hello World!")

' Opdracht 1B: Hello World 2
' Sla de tekst uit de vorige opdracht op in een variable en geef de inhoud van de variable weer in een Message Box
Dim helloWorld
helloWorld = "Hello World"

MsgBox(helloWorld)

' Opdracht 1C: Hello World 3
' Vraag de gebruiker met een InputBox om zijn/haar naam en sla dat op in een variable. Geef daarna die waarde weer
' in een Message Box met daarin de tekst uit de vorige opdracht en de naam daarachter geplakt.
Dim userName
userName = InputBox("Wat is je naam?", "Naam")

MsgBox(helloWorld & ", " & userName & "!")

' Opdracht 2A: Rekenmachine
' Vraag de gebruiker wat de uitkomst van  24 * 16 is. Laat de gebruiker daarna weten dat het antwoord juist is, of dat het
' antwoord onjuist is. Bonuspunten als je het antwoord van de gebruiker verwerkt in het bericht in de Message Box.
' Tip: Met de functie CInt zet je een string om naar een integer (maw een stukje tekst naar een getal zonder decimalen)
Dim correctAnswer
correctAnswer = 24 * 16

Dim userAnswer
userAnswer = InputBox("Wat is 24 * 16?")

If CInt(userAnswer) = correctAnswer Then
  MsgBox(userAnswer & " is het juiste antwoord")
Else
  MsgBox(userAnswer & " is onjuist. Het juiste antwoord is " & correctAnswer)
End If

' Opdracht 2B: Rekenmachine
' Kies een getal, sla dat op in een variable en geeft dat weer in een InputBox. Vraag de gebruiker een getal te kiezen 
' dat groter is dan jouw getal. Laat de gebruiker weten of het antwoord juist is.
Dim chosenNumber
chosenNumber = 42

userAnswer = InputBox("Kies een getal dat groter is dan " & chosenNumber)

If CInt(userAnswer) > chosenNumber Then
  MsgBox("Great Success!")
Else
  MsgBox("Great Failure!")
End if

' Opdracht 2C: Rekenmachine
' Kies opnieuw een getal en vraag de gebruiker weer een getal te kiezen. Geef weer of dat getal tussen de getallen die je 
' gekozen hebt in ligt. Laat de gebruiker weer weten of het antwoord juist is.
Dim anotherChosenNumber
anotherChosenNumber = 1899

userAnswer = InputBox("Kies een getal dat tussen " & chosenNumber & " en " & anotherChosenNumber & " in ligt.")

If CInt(userAnswer) > chosenNumber AND CInt(userAnswer) < anotherChosenNumber Then
  MsgBox("More Great Success!")
Else
  MsgBox("Not So Great Success!")
End If

' Opdracht 2D: Rekenmachine
' Declareer een variable met een array en zet daar een de getallen 1 t/m 5 in weer in. Geef vervolgens het derde getal
' uit de array weer in een MessageBox
Dim takeFive(5)
takeFive(0) = 1
takeFive(1) = 2
takeFive(2) = 3
takeFive(3) = 4
takeFive(4) = 5

MsgBox(takeFive(2))

' Opdracht 2E: Rekenmachine
' Neem de inhoud van je array uit de vorige opdracht en tel de inhoud bij elkaar op. Geef het resultaat weer in een MessageBox.
' Controleer of de inhoud klopt.
Dim sumResult
For Each val In takeFive
  sumResult = sumResult + val
Next

MsgBox(sumResult)

' Opdracht 2D: Rekenmachine
' Vraag de gebruiker om de uitkomst van 3 ^ 3. Houd niet op met vragen totdat de gebruiker het juiste antwoord geeft.
Dim answerIsCorrect
answerIsCorrect = False

Do
  userAnswer = InputBox("Wat is 3 ^ 3?", "Take 1")
  answerIsCorrect = CInt(userAnswer) = 9
Loop Until answerIsCorrect

' Opdracht 3A: Subs en Functions
' Verwerk de vorige opdracht in een Sub. Roep de Sub daarna aan om de gebruiker nog een keer om het antwoord te vragen.
Sub keepNagging()
  Dim answerIsCorrect
  answerIsCorrect = False

  Do
    userAnswer = InputBox("Wat is 3 ^ 3?", "Take 2")
    answerIsCorrect = CInt(userAnswer) = 9
  Loop Until answerIsCorrect
End Sub

Call keepNagging()

' Opdracht 3B: Subs en Functions
' Doe hetzelfde als in de vorige opdracht, maar parametriseer het bericht in de InputBox deze keer bij het aanroepen 
' van de Sub.
Sub nagResiliently(message)
  Dim answerIsCorrect
  answerIsCorrect = False

  Do
    userAnswer = InputBox(message, "Take 3")
    answerIsCorrect = CInt(userAnswer) = 9
  Loop Until answerIsCorrect
End Sub

Call nagResiliently("Wat is 3 ^ 3?")

' Opdracht 3C: Subs en Functions
' Schrijf een functie die twee getallen bij elkaar optelt en daarna het kwadraat daarvan uitrekent. Roep de function 
' aan en geef het resultaat weer in een MessageBox.
Function workMagic(number1, number2)
  Dim sum
  sum = number1 + number2
  workMagic = sum * sum
End Function

MsgBox(workMagic(3, 4))
