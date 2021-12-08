# Handige webpagina's:
  - https://www.tutorialspoint.com/vbscript/index.htm
  - https://www.tutorialspoint.com/vbscript/vbscript_dialog_boxes.htm
  - https://www.w3schools.com/asp/asp_ref_vbscript_functions.asp

# Wat is VBScript?
- Voluit: Microsoft Visual Basic Scripting Edition.
- Onderdeel van een familie van programmeertalen van MS: VBA, VB.NET en VBScript.

Een scripting language gebaseerd op Visual Basic die kan worden gebruikt om allerlei repetetieve taken in Windows te automatiseren.
Voorheen kon het ook worden gebruikt in Internet Explorer (maar niemand deed dat, want JavaScript). Microsoft beschouwt VBScript als "legacy",
wat betekent dat het niet meer actief wordt doorontwikkeld en op termijn uitgefaseerd gaat worden (maar dat kan heel lang duren).

# Variables
Variables declareer je voordat je ze gebruikt (hoeft niet altijd, maar doe maar wel):
  
  Dim variableName, anotherVariableName

Assignment gebeurt zo:
  
  variableName = "Een stukje tekst"
  variableName = 42
  variableName = 42.5

# Operators
Variables kan je gebruiken met verschillende operators. Er zijn verschillende categorieen van operators:

## Arithmetic operators:
  - +
  - -
  - *
  - /
  - %
  - ^

## Comparison operators
  - =
  - <>
  - \>
  - <
  - \>=
  - <=>

## Logical operators
  - AND
  - OR
  - NOT
  - XOR

# Concatenation operators
  - &

# Arrays
Arrays zijn een geavanceerder soort variable, die een verzameling van waardes bewaren (ipv 1 losse waarde). De items in een array hebben een index, die bij 0 begint.
Je declareert ze als volgt:

  Dim aNiceArrayName()
  Dim(0) = "Dit is een waarde"
  Dim(1) = 42
  anotherVariable = aNiceArrayName(0)

# Decision making
Is de afhandeling van je script afhankelijk van bepaalde voorwaardes? Daarvoor gebruik je If en Else:

  If [boolean expression] Then
    ' Doe iets hier
  End If
  If [boolean] Then
    ' Doe iets
  Else
    ' Doe iets anders
  End If

# Loops
Wil je een actie uitvoeren over alle items in een array? Dan kan je daarvoor een loop gebruiken. Er zijn een aantal verschillende soorten,
maar de meest gebruikte moet de For Each zijn:

For Each thingy In aNiceArray
  ' Do something
Next

Do
  ' Do something
Loop Until [boolean]

# Functions
Een klein herbruikbaar stukje code dat je zo vaak kan aanroepen als je wilt. Het helpt ook om je script overzichtelijk te houden.
Het idee is exact hetzelfde als met Excel formules en functions in SQL. Een function heeft een naam en je stopt er data in.
Eventueel geeft het je daarna weer data terug.

Function functieNaam(parameter1, parameter2)
  ' Do something
  ' Return something
  functieNaam = [resultaat]
End Function

Dim output
output = functieNaam(1, 2)

# Sub procedures
Ook een klein herbruikbaar stukje code, maar waar een functie je iets een nieuwe waarde teruggeeft, voert een sub alleen maar een taak uit (en geeft dus niets terug)
Denk hierbij aan het wegschrijven van stuk tekst naar een bestand. Subs moet je aanroepen met Call, anders gaat vbscript klagen

Sub subNaam(parameter1, parameter2, parameter3)
  ' Doe iets
End Sub

Call subNaam(1, 2, 3)
