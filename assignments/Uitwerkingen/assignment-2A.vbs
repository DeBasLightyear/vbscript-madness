' ###############################################
'               HIER NIET AANKOMEN!!
' De opdracht begint verder onderaan dit bestand.
' ###############################################
Function parseStringToArray(toParse)
    Dim preResult
    preResult = toParse

    ' remove all spaces and new lines
    preResult = Replace(preResult, " ", "", 1, -1, 1)
    preResult = Replace(preResult, vbLf, "", 1, -1, 0)
    preResult = Replace(preResult, vbCrLf, "", 1, -1, 0)

    ' remove trailing square bracket
    preResult = Left(preResult, Len(preResult) - 1)

    ' remove opening square brackets
    preResult = Replace(preResult, "[", "", 1, -1, 1)

    ' split outer array
    result = Split(preResult, "],")

    ' split inner arrays
    For i = 0 to UBound(result)
        result(i) = Split(result(i), ",")
    Next

    ' return result
    parseStringToArray = result
End Function

Function executeQueryServer(sql)
    ' Docs: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms760305(v=vs.85)
    ' Create an http object
    Dim xmlHttpReq
    Set xmlHttpReq = CreateObject("MSXML2.XMLHTTP.6.0")

    ' Set up the request
    Call xmlHttpReq.open("POST", "https://a8a9-195-78-54-218.ngrok.io/dvdrental/", false)
    Call xmlHttpReq.setRequestHeader("Content-Type", "application/json")

    ' Remove carriage return character (since UNIX servers don't like it) and fire the SQL
    dim cleanSQl
    cleanSQl = Replace(sql, vbCrLf, " ", 1, -1, 0)

    Call xmlHttpReq.send("{""sql"": """ & cleanSQl & """}")

    executeQueryServer = parseStringToArray(xmlHttpReq.responseText)
End Function

Function readTextFile(pathToFile)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set file = fso.OpenTextFile(pathToFile, 1)
    readTextFile = file.ReadAll()
End Function

Sub writeTextToFile(content, fileName)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim filePath
    filePath = "./" & fileName
    
    ' open a new file for writing and add the content
    Set file = fso.OpenTextFile(filePath, 2, True)
    file.Write(content)
    file.Close()
End Sub

' ###############################################
' ###############################################


' ###################################################################################
'                                        FINALE
' ###################################################################################
' Voor de laatste opdracht gaan we het uitvoeren van een query en het exporteren van 
' het resultaat automatiseren. Hiervoor zal je dus een query schrijven en die query
' geautomatiseerd uitvoeren, zodat je daarna het resultaat kan wegschrijven.
' Hier voor moet je de volgende stappen doorlopen:
'   1. Schrijf een query die aan de volgende eisen voldoet:
'       A. Koppel de tabel "film" aan de tabel "customer";
'       B. Reken uit wat iedere klant in totaal heeft uitgegeven aan gehuurde films (hint: aggregate function);
'       C. Order dit op zo'n manier dat je een ranglijst krijgt van groot naar klein;
'       D. Zorg ervoor dat de voornaam en achternaam van de klant in de output staat.
'   2. Sla die op in een .sql bestand;
'   3. Schrijf daarna een VB Script dat:
'       A. Het .sql bestand opent dmv de readTextFile functie;
'       B. De query vervolgens uitvoert dmv de executeQueryServer functie;
'       C. Het resultaat omzet naar een csv string;
'       D. De csv string wegschrijft naar een .csv bestand dmv de writeTextToFile functie;
'       E. Laat aan de gebruiker weten dat er een bestand is geschreven;
'   4. Open het bestand in Excel en aanschouw het resultaat.

' Stap 3A: Het SQL-bestand lezen
Dim sql
sql = readTextFile("./assignment-2A.sql")

' Stap 3B: De query uitvoeren
Dim queryResult
queryResult = executeQueryServer(sql)

' Stap 3C: Het resultaat omzetten naar een csv-string
Dim csvString
For Each row in queryResult
    ' join all row items together on a semicolon and add a new line
    csvString = csvString & Join(row, ";") & vbLf
Next

' Stap 3D: De csv-string naar een tekstbestand wegschrijven
Call writeTextToFile(csvString, "assignment-2A.csv")

' Stap 3E: De gebruiker laten weten dat er een bestand is geschreven
MsgBox("Er zijn dingen gebeurd.")
