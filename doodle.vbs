Function readTextFile(pathToFile)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set file = fso.OpenTextFile(pathToFile, 1)
    readTextFile = file.ReadAll()
End Function

Function parseStringToArray(toParse)
    Dim preResult
    Dim result()
    preResult = toParse

    ' remove all spaces and new lines
    preResult = Replace(preResult, " ", "", 1, -1, 1)
    preResult = Replace(preResult, vbLf, "", 1, -1, 0)
    preResult = Replace(preResult, vbCrLf, "", 1, -1, 0)
    MsgBox(preResult)

    ' remove trailing square bracket
    preResult = Left(preResult, Len(preResult) - 1)
    MsgBox(preResult)

    ' remove opening square brackets
    preResult = Replace(preResult, "[", "", 1, -1, 1)
    MsgBox(preResult)

    ' split outer array
    preResult = Split(preResult, "],")

    ' split inner arrays and return
    For Each item in preResult
        result(Len(result) + 1) = Split(item, ",")
    Next

    parseStringToArray = result
End Function

Dim doodle
doodle = "./doodle.txt"

Dim txt
txt = readTextFile(doodle)

MsgBox("[[Foo, Bar, Wololo],[1, 2, 3],[4, 5, 6],[7, 8, 9],]")

Dim res
res = parseStringToArray(txt)

MsgBox("Great Success?")
